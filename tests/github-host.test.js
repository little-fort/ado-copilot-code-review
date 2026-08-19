/**
 * Tests for the GitHub Enterprise host normalizer in index.ts (issue #59).
 *
 * Requires the compiled task (`npm run build`) — CI builds before testing.
 * The value produced here is exported as GH_HOST, which expects a bare
 * hostname, while users commonly paste full URLs.
 */

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

describe('normalizeGitHubHost (issue #59)', () => {
    let normalizeGitHubHost;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        ({ normalizeGitHubHost } = require(compiledTask));
    });

    it('passes a bare hostname through unchanged', () => {
        assert.equal(normalizeGitHubHost('subdomain.ghe.com'), 'subdomain.ghe.com');
    });

    it('strips an https scheme', () => {
        assert.equal(normalizeGitHubHost('https://subdomain.ghe.com'), 'subdomain.ghe.com');
    });

    it('strips an http scheme case-insensitively', () => {
        assert.equal(normalizeGitHubHost('HTTP://subdomain.ghe.com'), 'subdomain.ghe.com');
    });

    it('strips trailing slashes', () => {
        assert.equal(normalizeGitHubHost('https://subdomain.ghe.com/'), 'subdomain.ghe.com');
        assert.equal(normalizeGitHubHost('subdomain.ghe.com//'), 'subdomain.ghe.com');
    });

    it('trims surrounding whitespace', () => {
        assert.equal(normalizeGitHubHost('  subdomain.ghe.com '), 'subdomain.ghe.com');
    });
});
