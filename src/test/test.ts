import * as chai from 'chai';
import * as chaiAsPromised from 'chai-as-promised';
import * as fs from 'fs';
import * as request from 'request';
import * as rp from 'request-promise';

chai.use(chaiAsPromised);
chai.should();
let expect = chai.expect;
let baseUri = 'https://verificationservice.osi.office.net/ova/addincheckingagent.svc/api/addincheck?lang=';
let options = {
  uri: baseUri,
  method: 'POST',
  headers: {
    'Content-Type': 'application/xml'
  },
  resolveWithFullResponse: true
};

function callOmexService (file, options) {
  let fileStream = fs.createReadStream(file);
  return fileStream.pipe(rp(options))
    .then((response) => { return response.statusCode; })
    .catch((err) => { throw err.statusCode; });
}

describe('Test manifest files', () => {
  describe('Valid - 200', () => {
    let result = '';

    it('should return validation passed with code 200', () => {
      let manifest = './manifest-to-test/valid_excel.xml';
      return callOmexService(manifest, options).should.eventually.equal(200);
    });
  });

  describe('Invalid - 400', () => {
    let result = '';

    it('should return validation failed with code 400', () => {
      let manifest = './manifest-to-test/invalid_400.xml';
      return callOmexService(manifest, options).should.eventually.throw;
      ;
    });
  });



});
