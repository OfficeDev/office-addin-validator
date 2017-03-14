import * as chai from 'chai';
import { callOmexService } from '../index';

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

describe('Test Valid Manifest Files', () => {
  describe('Excel', () => {
    let result = '';
    before((done) => {
      let manifest = './manifest-to-test/valid_excel.xml';
      callOmexService(manifest, options).then((response) => {
        result = response.statusCode;
        done();
      });
    });

    it('should return validation passed', () => {
      expect(result).to.equal(200);
    });
  });

});
