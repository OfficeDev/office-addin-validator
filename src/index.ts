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

let baseUri = 'https://verificationservice.osi.office.net/ova/addincheckingagent.svc/api/addincheck?lang=';
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
  .option('-l, --language', 'localization language', 'en-US')
  .action(async (manifest) => {
    // progress bar start
    status.start({
      pattern: '    {uptime.green} {spinner.dots.green} Calling validation service...'
    });
    // set localization parameter
    let language = commander.language;
    options.uri = baseUri + language;
    try {
      let response = await callOmexService(manifest, options);
      if (response.statusCode === 200) {
        let formattedBody = JSON.parse(response.body.trim());
        let validationReport = formattedBody.checkReport.validationReport;
        let validationResult = validationReport.result;
        let validationErrors = [];
        let validationWarnings = [];
        let validationInfos = [];

        getNestedObj(validationReport, 'errors', validationErrors);
        getNestedObj(validationReport, 'warnings', validationWarnings);
        getNestedObj(validationReport, 'infos', validationInfos);

        console.log('----------------------');
        if (validationResult === 'Passed') {
          // supported products only exist when manifest is valid
          let supportedProducts = formattedBody.checkReport.details.supportedProducts;
          console.log(`${chalk.bold('Validation: ')}${chalk.bold.green('Passed')}`);
          logValidationReport(validationWarnings, 'warning');
          logValidationReport(validationInfos, 'info');
          logSupportedProduct(supportedProducts);
        } else {
          console.log(`${chalk.bold('Validation: ')}${chalk.bold.red('Failed')}`);
          logValidationReport(validationErrors, 'error');
          logValidationReport(validationWarnings, 'warning');
          logValidationReport(validationInfos, 'info');
        }
        console.log('----------------------');
      } else {
        console.log('Unexpected Error');
      }
    }
    catch (err) {
      let statusCode = err['statusCode'];
      logError(statusCode);
      // exit node process when error is thrown
      process.exitCode = 1;
    }
    // stop progress bar
    finally { status.stop(); }
  })
  .parse(process.argv);

async function callOmexService(file, options) {
  let fileStream = fs.createReadStream(file);
  let response = await fileStream.pipe(rp(options));
  return response;
}

function logError(statusCode) {
  console.log('----------------------');
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
  console.log('----------------------');
}

function getNestedObj(obj, item, result) {
  for (let i = 0; i < obj[item].length; i++) {
    let itemTitle = obj[item][i].title;
    let itemDetail = obj[item][i].detail;
    let itemLink = obj[item][i].link;
    let itemCollection = {
      'title': itemTitle,
      'detail': itemDetail,
      'link': itemLink
    };
    result.push(JSON.stringify(itemCollection));
  }

  return result;
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
    for (let i = 0; i < obj.length; i++) {
      let jsonObj = JSON.parse(obj[i]);
      console.log('  - ' + jsonObj.title + ': ' + jsonObj.detail + ' (link: ' + jsonObj.link + ')');
    }
  }
}

function logSupportedProduct(obj) {
  if (obj.length > 0) {
    console.log(`With this manifest, the store will test your add-in against the following platforms:`);
    for (let i = 0; i < obj.length; i++) {
      console.log('  - ' + obj[i].title + ', Version: ' + obj[i].version);
    }
  }
}
