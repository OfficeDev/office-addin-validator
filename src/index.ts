#!/usr/bin/env node

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as commander from 'commander';
import * as fs from 'fs';
import * as util from './util';
import * as appInsights from 'applicationinsights';

let insight = appInsights.getClient('78cc7757-c7a2-4382-b801-bce73cf33d7a');

commander
  .arguments('<manifest>')
  .action(async (manifest) => {
    if (fs.existsSync(manifest)) {
      process.exitCode = await util.validateManifest(manifest);
      process.exit();
    } else {
      console.log('-------------------------------------');
      console.log('Error: Please provide a valid local manifest file path.');
      console.log('-------------------------------------');
      insight.trackException(new Error('Manifest file path is not valid.'));
      // update node process exit code when file does not exit
      process.exitCode = 1;
    }
  }).parse(process.argv);
