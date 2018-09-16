import * as assert from 'assert';
import * as mocha from 'mocha';
import { readManifestFile } from '../src/manifest';


describe('Manifest', function() {
  describe('readManifestInfo', function() {
    it('should read the manifest info', async function() {
      const info = await readManifestFile('test/manifest.xml')

      assert.strictEqual(info.defaultLocale, 'en-US');
      assert.strictEqual(info.description, 'Describes this Office Add-in.');
      assert.strictEqual(info.displayName, 'Office Add-in Name');
      assert.strictEqual(info.id, '132a8a21-011a-4ceb-9336-6af8a276a288');
      assert.strictEqual(info.officeAppType, 'TaskPaneApp');
      assert.strictEqual(info.providerName, 'ProviderName');
      assert.strictEqual(info.version, '1.2.3.4');
    });
    it('should throw an error if there is a bad xml end tag', async function() {  
        let result;
        try {
          const info = await readManifestFile('test/manifest.incorrect-end-tag.xml');
        } catch (err) {          
          result = err;
        };
        assert.equal(result, "Unable to parse the manifest file: test/manifest.incorrect-end-tag.xml. \nError: Unexpected close tag\nLine: 8\nColumn: 46\nChar: >");        
    });
    it('should handle a missing description', async function() {
      const info = await readManifestFile('test/manifest.no-description.xml')

      assert.strictEqual(info.defaultLocale, 'en-US');
      assert.strictEqual(info.description, undefined);
      assert.strictEqual(info.displayName, 'Office Add-in Name');
      assert.strictEqual(info.id, '132a8a21-011a-4ceb-9336-6af8a276a288');
      assert.strictEqual(info.officeAppType, 'TaskPaneApp');
      assert.strictEqual(info.providerName, 'ProviderName');
      assert.strictEqual(info.version, '1.2.3.4');
    });
  });
});