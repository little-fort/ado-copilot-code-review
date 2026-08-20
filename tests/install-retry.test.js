/**
 * Tests for the installer retry helper in index.ts (issue #57).
 *
 * Requires the compiled task (`npm run build`) because the helper is imported
 * from CopilotCodeReviewV1/index.js — CI builds before running tests. The
 * module's run() only fires when index.js is the process entry point, so
 * importing it here is side-effect free.
 */

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

describe('runInstallerWithRetry (issue #57)', function () {
    // Real subprocess spawns plus a 2s backoff make these slower than unit tests
    this.timeout(30000);

    let tmpDir;
    let helperScript;
    let runInstallerWithRetry;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        ({ runInstallerWithRetry } = require(compiledTask));

        tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'install-retry-'));
        // Helper child script: exits 1 until its counter file records
        // <failures> prior runs, then exits 0. Lets tests simulate an
        // installer that succeeds only after transient failures.
        helperScript = path.join(tmpDir, 'flaky.js');
        fs.writeFileSync(helperScript, `
            const fs = require('fs');
            const [counterFile, failures] = process.argv.slice(2);
            const n = fs.existsSync(counterFile) ? Number(fs.readFileSync(counterFile, 'utf8')) : 0;
            fs.writeFileSync(counterFile, String(n + 1));
            process.exit(n < Number(failures) ? 1 : 0);
        `);
    });

    after(() => {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    function flakyCommandArgs(name, failures) {
        const counterFile = path.join(tmpDir, `${name}.count`);
        // Quote paths: the helper spawns with shell:true and temp paths may contain spaces
        return ['node', [`"${helperScript}"`, `"${counterFile}"`, String(failures)], counterFile];
    }

    it('succeeds immediately when the installer succeeds on the first attempt', async () => {
        const [cmd, args, counterFile] = flakyCommandArgs('first-try', 0);
        await runInstallerWithRetry('Test installer', cmd, args);
        assert.equal(fs.readFileSync(counterFile, 'utf8'), '1', 'installer should have run exactly once');
    });

    it('retries a transient failure and succeeds on a later attempt', async () => {
        const [cmd, args, counterFile] = flakyCommandArgs('second-try', 1);
        await runInstallerWithRetry('Test installer', cmd, args, 3);
        assert.equal(fs.readFileSync(counterFile, 'utf8'), '2', 'installer should have run twice');
    });

    it('fails after exhausting all attempts, reporting the attempt count', async () => {
        const [cmd, args, counterFile] = flakyCommandArgs('always-fails', 99);
        await assert.rejects(
            () => runInstallerWithRetry('Test installer', cmd, args, 2),
            /Test installer failed after 2 attempts/);
        assert.equal(fs.readFileSync(counterFile, 'utf8'), '2', 'installer should have run exactly maxAttempts times');
    });
});
