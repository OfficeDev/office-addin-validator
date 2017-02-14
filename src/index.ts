#!/usr/bin/env node

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as commander from 'commander';

let curl = require('curlrequest');
let options = {
  url: 'https://verificationservice.osi.office.net/ova/addincheckingagent.svc/api/addincheck?lang=zh-CH',
  method: 'POST',
  headers: {
    'Content-Type': 'application/xml'
  },
  data: null,
  include: true
};

commander
  .arguments('<manifest>')
  .option('-l, --language [optional]', 'localization language')
  .action((manifest) => {
    options.data = '@' + manifest;

    curl.request(options, (err, data) => {
      console.log('processing');
      console.log(data);
      console.log(err);
    });
  })
  .parse(process.argv);


