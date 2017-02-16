#!/usr/bin/env node

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as commander from 'commander';
import * as prettyjson from 'prettyjson';
import * as fs from 'fs';
import * as request from 'request';
import * as chalk from 'chalk';


let options = {
  uri: 'https://verificationservice.osi.office.net/ova/addincheckingagent.svc/api/addincheck?lang=zh-CH',
  method: 'POST',
  headers: {
    'Content-Type': 'application/xml'
  }
};


commander
  .arguments('<manifest>')
  .option('-l, --language [optional]', 'localization language')
  .action((manifest) => {
    callOmexService(manifest, options, (formattedBody) => {
        let validationReport = formattedBody.checkReport.validationReport;
        let validationResult = validationReport.result;
        let validationErrors = [];
        let validationWarnings = [];
        let validationInfos = [];

        getNestedObj(validationReport, 'errors', validationErrors);
        getNestedObj(validationReport, 'warnings', validationWarnings);
        getNestedObj(validationReport, 'infos', validationInfos);

        if (validationResult === 'Passed') {
          // supported products only exist when manifest is valid
          let supportedProducts = formattedBody.checkReport.details.supportedProducts;

          console.log(`${chalk.bold('Validation: ')}${chalk.bold.green('Passed')}`);
          console.log(`\nWarning(s):`);
          logValidationReport(validationWarnings, 'yellow');
          console.log(`\nAdditional Information:`);
          logValidationReport(validationInfos, '');
          console.log(`\nWith this manifest, the store will test your add-in against the following platforms:`);
          logSupportedProduct(supportedProducts);
        } else {
          console.log(`${chalk.bold('Validation: ')}${chalk.bold.red('Failed')}`);
          console.log(`\nErrors(s):`);
          logValidationReport(validationErrors, 'red');
          console.log(`\nWarning(s):`);
          logValidationReport(validationWarnings, 'yellow');
          console.log(`\nAdditional Information:`);
          logValidationReport(validationInfos, 'dim');
        }
    });
  })
  .parse(process.argv);

function callOmexService (file, options, callback) {
  let formattedBody = {};
  fs.createReadStream(file)
    .pipe(request(options, (err, res, body) => {
      formattedBody = JSON.parse(body.trim());

      return callback(formattedBody);
    }));
}

function getNestedObj (obj, item, result) {
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

function logValidationReport (obj, color) {
  if (obj.length > 0) {
    for (let i = 0; i < obj.length; i++) {
      let jsonObj = JSON.parse(obj[i]);
      console.log(`${chalk[color](jsonObj.title + ': ')}` + jsonObj['detail'] + ' (link: ' + jsonObj['link'] + ')');
    }
  } else {
    // TODO: get language
    console.log('N/A');
  }
}

function logSupportedProduct (obj) {
  if (obj.length > 0) {
    for (let i = 0; i < obj.length; i++) {
      console.log(obj[i].title + ', Version: ' + obj[i].version);
    }
  } else {
    // TODO; get language
    console.log('N/A');
  }
}
