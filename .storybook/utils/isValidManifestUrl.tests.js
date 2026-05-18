import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { isValidManifestUrl, TRUSTED_PREFIX } from './isValidManifestUrl.js';

describe('isValidManifestUrl', () => {
  describe('valid URLs', () => {
    it('should accept a URL directly under the trusted prefix', () => {
      const url = `${TRUSTED_PREFIX}samples/my-sample/manifest.json`;
      assert.equal(isValidManifestUrl(url), true);
    });

    it('should accept a URL in a nested path under the trusted prefix', () => {
      const url = `${TRUSTED_PREFIX}samples/app/deep/path/manifest.json`;
      assert.equal(isValidManifestUrl(url), true);
    });

    it('should accept the trusted prefix itself', () => {
      assert.equal(isValidManifestUrl(TRUSTED_PREFIX), true);
    });
  });

  describe('path traversal attacks', () => {
    it('should reject a URL with ../ that traverses to another repository', () => {
      const url = 'https://raw.githubusercontent.com/pnp/mgt-samples/main/../../../attacker/repo/main/manifest.json';
      assert.equal(isValidManifestUrl(url), false);
    });

    it('should reject a URL with encoded traversal segments (%2e%2e)', () => {
      const url = 'https://raw.githubusercontent.com/pnp/mgt-samples/main/%2e%2e/%2e%2e/%2e%2e/attacker/repo/main/evil.js';
      assert.equal(isValidManifestUrl(url), false);
    });

    it('should reject a URL that uses ../ to escape to a sibling repo', () => {
      const url = 'https://raw.githubusercontent.com/pnp/mgt-samples/main/../other-repo/main/payload.json';
      assert.equal(isValidManifestUrl(url), false);
    });

    it('should reject a URL with multiple ../ segments', () => {
      const url = 'https://raw.githubusercontent.com/pnp/mgt-samples/main/../../evil-org/evil-repo/main/manifest.json';
      assert.equal(isValidManifestUrl(url), false);
    });
  });

  describe('different origins and invalid URLs', () => {
    it('should reject a URL from a completely different domain', () => {
      const url = 'https://evil.com/pnp/mgt-samples/main/manifest.json';
      assert.equal(isValidManifestUrl(url), false);
    });

    it('should reject a URL from a different GitHub user/org', () => {
      const url = 'https://raw.githubusercontent.com/attacker/malicious-repo/main/manifest.json';
      assert.equal(isValidManifestUrl(url), false);
    });

    it('should reject a URL that only partially matches the prefix', () => {
      const url = 'https://raw.githubusercontent.com/pnp/mgt-samples-evil/main/manifest.json';
      assert.equal(isValidManifestUrl(url), false);
    });

    it('should reject a data: URI', () => {
      const url = 'data:application/json;base64,W3sicHJldmlldyI6eyJqcyI6Imh0dHA6Ly9ldmlsLmNvbS9ldmlsLmpzIn19XQ==';
      assert.equal(isValidManifestUrl(url), false);
    });

    it('should reject a javascript: URI', () => {
      const url = 'javascript:alert(1)';
      assert.equal(isValidManifestUrl(url), false);
    });
  });

  describe('null/empty/malformed inputs', () => {
    it('should reject null', () => {
      assert.equal(isValidManifestUrl(null), false);
    });

    it('should reject undefined', () => {
      assert.equal(isValidManifestUrl(undefined), false);
    });

    it('should reject empty string', () => {
      assert.equal(isValidManifestUrl(''), false);
    });

    it('should reject a malformed URL', () => {
      assert.equal(isValidManifestUrl('not-a-url'), false);
    });

    it('should reject a relative path', () => {
      assert.equal(isValidManifestUrl('../../../etc/passwd'), false);
    });
  });
});
