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

        getNestedObj(validationReport, 'errors', validationErrors);
        getNestedObj(validationReport, 'warnings', validationWarnings);

        console.log(validationErrors);
        console.log(validationWarnings);
        console.log(formattedBody);

        if (validationResult = 'Passed') {
          console.log(`Validation: ${chalk.green('Passed')}`);
        } else {
          console.log(`Validation: ${chalk.red('Failed')}`);
        }
    });

        // console.log(prettyjson.render(formattedBody, {
        //   keysColor: 'rainbow',
        //   dashColor: 'magenta',
        //   stringColor: 'white'
        // }));
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
    result.push(JSON.stringify(itemTitle));
  }

  return result;
}
