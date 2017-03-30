#!/usr/bin/env node

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as commander from 'commander';
import * as fs from 'fs';
import * as request from 'request';
import * as rp from 'request-promise';
import * as chalk from 'chalk';
import * as status from 'node-status';
import * as appInsights from 'applicationinsights';

let insight = appInsights.getClient('78cc7757-c7a2-4382-b801-bce73cf33d7a');
let baseUri = 'https://verificationservice.osi.office.net/ova/addincheckingagent.svc/api/addincheck';
let options = {
  uri: baseUri,
  method: 'POST',
  headers: {
    'Content-Type': 'application/xml'
  },
  resolveWithFullResponse: true
};

commander
  .arguments('<manifest>')
  .action(async (manifest) => {
    try {
      if (fs.existsSync(manifest)) {
        // progress bar start
        status.start({
          pattern: '  {uptime.green} {spinner.dots.green} Calling validation service...'
        });
        let response = await callOmexService(manifest, options);
        if (response.statusCode === 200) {
          let formattedBody = JSON.parse(response.body.trim());
          let validationReport = formattedBody.checkReport.validationReport;
          let validationResult = validationReport.result;

          console.log('\n-------------------------------------');
          switch (validationResult) {
            case 'Passed':
              // supported products only exist when manifest is valid
              let supportedProducts = formattedBody.checkReport.details.supportedProducts;
              console.log(`${chalk.bold('Validation: ')}${chalk.bold.green('Passed')}`);
              logValidationReport(validationReport.warnings, 'warning');
              logValidationReport(validationReport.infos, 'info');
              logSupportedProduct(supportedProducts);
              break;
            case 'Failed':
              console.log(`${chalk.bold('Validation: ')}${chalk.bold.red('Failed')}`);
              logValidationReport(validationReport.errors, 'error');
              logValidationReport(validationReport.warnings, 'warning');
              logValidationReport(validationReport.infos, 'info');
              break;
          }
          console.log('-------------------------------------');
          insight.trackEvent('Validation Results', { result: validationResult });
        } else {
          console.log('Unexpected program error.');
          insight.trackException(new Error('Unexpected program error.'));
        }
      } else {
        console.log('Error: Please provide a valid local manifest file path.');
        insight.trackException(new Error('Manifest file path is not valid.'));
        // exit node process when file does not exit
        process.exitCode = 1;
      }
    } catch (err) {
      let statusCode = err['statusCode'];
      logError(statusCode);
      insight.trackException(new Error('Service Error. Error Code: ' + statusCode));
      // exit node process when error is thrown
      process.exitCode = 1;
    } finally {
      // stop progress bar
      status.stop();
      process.exit();
    }
  }).parse(process.argv);

async function callOmexService(file, options) {
  let fileStream = fs.createReadStream(file);
  let response = await fileStream.pipe(rp(options));
  return response;
}

function logError(statusCode) {
  console.log('-------------------------------------');
  console.log(`${chalk.bold('Validation: ')}${chalk.bold.red('Failed')}`);
  console.log('  Error Code: ' + statusCode);
  console.log(`  ${chalk.bold.red('Error(s): ')}`);
  switch (statusCode) {
    case 400:
      console.log('  Request body does not contain a valid XML document, and/or is too large (capped at 256kb).');
      break;
    case 415:
      console.log('  Content-Type is not set to application/xml.');
      break;
    case 500:
      console.log('  Unexpected error.');
      break;
    case 503:
      console.log('  Service unavailable; API processing has been disabled via BRS.');
      break;
  }
  console.log('-------------------------------------');
}

function logValidationReport(obj, name) {
  if (obj.length > 0) {
    switch (name) {
      case 'error':
        console.log(`  ${chalk.bold.red('Error(s): ')}`);
        break;
      case 'warning':
        console.log(`  ${chalk.bold.yellow('Warning(s): ')}`);
        break;
      case 'info':
        console.log(`  Additional Information:`);
        break;
    }
    for (let i of obj) {
      console.log('  - ' + i.title + ': ' + i.detail + ' (link: ' + i.link + ')');
    }
  }
}

function logSupportedProduct(obj) {
  if (obj.length > 0) {
    let product = [];
    getProduct(obj, product);
    let unique = [...new Set(product)];
    console.log(`These Office applications will accept your add-in manifest:`);
    for (let i of unique) {
      console.log('  - ' + i);
    }
    console.log(`Based on the features and requirement sets implemented by your add-in, use the Add-in Availability Page (https://dev.office.com/add-in-availability) to determine which platforms to test your add-in against.`);
  }
}

function getProduct(array, result) {
  for (let i of array) {
    let itemTitle = i.title;
    result.push(itemTitle);
  }
}
