/**
 * Tests for the per-task folders in the CopilotCodeReview artifact (publishPromptArtifacts).
 *
 * Several review tasks in one job publish to the same artifact. Each task stages its files
 * under <stagingRoot>/<agent>-<key4>/ and uploads stagingRoot as one directory; the agent
 * keeps the path below the uploaded directory (FileContainerServer: containerPath + "/" +
 * the file's path relative to the source folder), so the files land in
 * CopilotCodeReview/<agent>-<key4>/ instead of overwriting each other at the artifact root.
 *
 * Requires the compiled task (`npm run build`); importing index.js is side-effect free.
 */

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

describe('prompt artifact folders', () => {
    let task;
    let tmpDir;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        task = require(compiledTask);
        tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'artifacts-'));
    });

    after(() => {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    // Lists every file under root as a forward-slash path relative to root
    function listRelative(root) {
        return fs.readdirSync(root, { recursive: true, withFileTypes: true })
            .filter(entry => entry.isFile())
            .map(entry => path.relative(root, path.join(entry.parentPath ?? entry.path, entry.name)).split(path.sep).join('/'))
            .sort();
    }

    it('names the folder after the agent and the first four hex digits of the review-state key', () => {
        const key = 'CopilotCodeReview.v1.5ab77b1ae2b8b58b';
        assert.equal(task.artifactSubfolderName('copilot', key), 'copilot-5ab7');
        assert.equal(task.artifactSubfolderName('claude', 'CopilotCodeReview.v1.06299b49eeabcc65'), 'claude-0629');
    });

    it('takes the digits from the key computeReviewStateKeyPrefix actually produces', () => {
        const key = task.computeReviewStateKeyPrefix({
            agent: 'copilot', model: 'gpt-6.1-sol', reasoningEffort: '', maxTurns: '', maxBudget: '',
            providerType: '', providerBaseUrl: '', diffOnly: false, promptMode: 'default', promptText: '',
            templateText: 'template', resolveThreads: 'all', suppressPositiveFeedback: false,
            maxMinorIssues: '5', workItemSource: 'none', ignoreCustomInstructions: false
        });
        const hash = key.split('.').pop();
        assert.match(hash, /^[0-9a-f]{16}$/);
        assert.equal(task.artifactSubfolderName('copilot', key), `copilot-${hash.slice(0, 4)}`);
    });

    it('stages the existing files under <stagingRoot>/<folder>/ and skips missing ones', () => {
        const source = path.join(tmpDir, 'checkout');
        fs.mkdirSync(source);
        fs.writeFileSync(path.join(source, 'PR_Details.txt'), 'details');
        fs.writeFileSync(path.join(source, '_claude_prompt.txt'), 'prompt');
        const stagingRoot = path.join(tmpDir, 'staging-1');

        const staged = task.stageArtifactFiles([
            path.join(source, 'PR_Details.txt'),
            path.join(source, 'Work_Item_Details.txt'), // not written this run
            undefined,
            path.join(source, '_claude_prompt.txt')
        ], stagingRoot, 'claude-0629');

        assert.deepEqual(staged, ['PR_Details.txt', '_claude_prompt.txt']);
        // These relative paths are what the agent appends to the CopilotCodeReview container folder
        assert.deepEqual(listRelative(stagingRoot), ['claude-0629/PR_Details.txt', 'claude-0629/_claude_prompt.txt']);
        assert.equal(fs.readFileSync(path.join(stagingRoot, 'claude-0629', 'PR_Details.txt'), 'utf8'), 'details');
    });

    it('copies the files, so later changes in the checkout do not reach the artifact', () => {
        const source = path.join(tmpDir, 'checkout-2');
        fs.mkdirSync(source);
        const file = path.join(source, 'Iteration_Details.txt');
        fs.writeFileSync(file, 'as given to the agent');
        const stagingRoot = path.join(tmpDir, 'staging-2');

        task.stageArtifactFiles([file], stagingRoot, 'copilot-5ab7');
        fs.writeFileSync(file, 'changed during the review');

        assert.equal(fs.readFileSync(path.join(stagingRoot, 'copilot-5ab7', 'Iteration_Details.txt'), 'utf8'), 'as given to the agent');
    });

    // Runs fn while capturing what it writes to stdout (where task-lib emits ##vso commands)
    function captureStdout(fn) {
        const original = process.stdout.write;
        let captured = '';
        process.stdout.write = (chunk, ...rest) => {
            captured += String(chunk);
            return true;
        };
        try {
            return { result: fn(), output: captured };
        } finally {
            process.stdout.write = original;
        }
    }

    it('uploads the staging folder once, into the shared CopilotCodeReview container folder', () => {
        const source = path.join(tmpDir, 'checkout-3');
        fs.mkdirSync(source);
        fs.writeFileSync(path.join(source, 'PR_Details.txt'), 'details');
        fs.writeFileSync(path.join(source, 'review-policy.json'), '{}');
        const stagingRoot = path.join(tmpDir, 'staging-4');

        const { result, output } = captureStdout(() => task.publishPromptArtifactFiles(
            [path.join(source, 'PR_Details.txt'), path.join(source, 'review-policy.json')], stagingRoot, 'copilot-5ab7'));

        assert.deepEqual(result, ['PR_Details.txt', 'review-policy.json']);
        const commands = output.split(/\r?\n/).filter(line => line.startsWith('##vso[artifact.upload'));
        // One command for the folder, not one per file, and the container folder is the
        // artifact root: the per-task folder comes from the path below stagingRoot
        assert.equal(commands.length, 1, output);
        assert.equal(commands[0], `##vso[artifact.upload containerfolder=CopilotCodeReview;artifactname=CopilotCodeReview;]${stagingRoot}`);
    });

    it('issues no upload when none of the files exist', () => {
        const stagingRoot = path.join(tmpDir, 'staging-5');
        const { result, output } = captureStdout(() => task.publishPromptArtifactFiles([path.join(tmpDir, 'missing.txt')], stagingRoot, 'claude-0629'));
        assert.deepEqual(result, []);
        assert.doesNotMatch(output, /artifact\.upload/);
    });

    it('creates nothing when there is nothing to publish', () => {
        const stagingRoot = path.join(tmpDir, 'staging-3');
        const staged = task.stageArtifactFiles([undefined, path.join(tmpDir, 'missing.txt')], stagingRoot, 'copilot-5ab7');
        assert.deepEqual(staged, []);
        // An empty upload folder would make the agent warn ("directory is empty"), so the task skips the upload
        assert.equal(fs.existsSync(stagingRoot), false);
    });
});
