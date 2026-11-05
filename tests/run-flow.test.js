/**
 * End-to-end tests of the task's run() flow up to the agent step (issue #67).
 *
 * The compiled task runs as a child process, configured through the INPUT_*
 * environment variables azure-pipelines-task-lib reads, against a local HTTP
 * server standing in for the Azure DevOps REST API. The real PowerShell scripts
 * run unmodified. No CLI agent is installed or run: one test ends at the
 * already-reviewed skip, the other at a deliberate prompt configuration error
 * (which, after the reorder, must stop the run before the CLI install step).
 *
 * Requires the compiled task (`npm run build`) and pwsh on PATH.
 */

const assert = require('node:assert/strict');
const child_process = require('node:child_process');
const fs = require('node:fs');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');

const taskDir = path.join(__dirname, '..', 'CopilotCodeReviewV1');
const compiledTask = path.join(taskDir, 'index.js');

const SOURCE_SHA = 'c'.repeat(40);
const TARGET_SHA = '1'.repeat(40);
const BASE_SHA = 'e'.repeat(40);

/**
 * Serves the Azure DevOps endpoints the fetch scripts call, keyed by path
 * (case-insensitive, query ignored). Unknown paths return 404.
 */
function startMockAzureDevOps(routes) {
    const requests = [];
    const server = http.createServer((req, res) => {
        const route = decodeURIComponent(req.url.split('?')[0]).toLowerCase();
        requests.push({ method: req.method, path: route });
        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', () => {
            const handler = Object.entries(routes).find(([suffix]) => route.endsWith(suffix.toLowerCase()));
            const [status, payload] = handler ? handler[1] : [404, { message: 'mock: no route' }];
            res.writeHead(status, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(payload));
        });
    });
    return new Promise(resolve => {
        server.listen(0, '127.0.0.1', () => resolve({ server, requests, port: server.address().port }));
    });
}

/** Default responses for PR #5 with a single iteration. */
function baseRoutes(overrides = {}) {
    const pr = '/testorg/TestProject/_apis/git/repositories/TestRepo/pullrequests/5';
    return {
        [`${pr}/iterations/1/changes`]: [200, { changeEntries: [{ changeType: 'edit', item: { path: '/src/a.cs' } }], nextSkip: 0, nextTop: 0 }],
        [`${pr}/iterations`]: [200, { count: 1, value: [{ id: 1, reason: 'create', sourceRefCommit: { commitId: SOURCE_SHA }, targetRefCommit: { commitId: TARGET_SHA }, commonRefCommit: { commitId: BASE_SHA } }] }],
        [`${pr}/commits`]: [200, { count: 1, value: [{ commitId: SOURCE_SHA, comment: 'Change a', author: { name: 'Dev', date: '2026-01-01T00:00:00Z' } }] }],
        [`${pr}/properties`]: [200, { count: 0, value: {} }],
        [`${pr}/workitems`]: [200, { count: 0, value: [] }],
        [`${pr}/threads`]: [200, { count: 0, value: [] }],
        [pr]: [200, {
            pullRequestId: 5, title: 'Flow test', description: '', status: 'active',
            repository: { name: 'TestRepo' }, sourceRefName: 'refs/heads/feature', targetRefName: 'refs/heads/main',
            isDraft: false, mergeStatus: 'succeeded', createdBy: { displayName: 'Dev', uniqueName: 'dev@example.com' },
            creationDate: '2026-01-01T00:00:00Z', reviewers: []
        }],
        ...overrides
    };
}

/** Runs the compiled task with the given inputs; resolves with its output. */
function runTask(port, workDir, tempDir, inputs = {}) {
    const env = { ...process.env };
    // Never inherit pipeline variables that would change the flow
    for (const key of Object.keys(env)) {
        if (/^(INPUT_|SYSTEM_PULLREQUEST_|BUILD_)/i.test(key)) {
            delete env[key];
        }
    }
    Object.assign(env, {
        INPUT_GITHUBPAT: 'test-github-token',
        INPUT_USESYSTEMACCESSTOKEN: 'false',
        INPUT_AZUREDEVOPSPAT: 'test-ado-token',
        INPUT_COLLECTIONURI: `http://127.0.0.1:${port}/testorg`,
        INPUT_PROJECT: 'TestProject',
        INPUT_REPOSITORY: 'TestRepo',
        INPUT_PULLREQUESTID: '5',
        SYSTEM_DEFAULTWORKINGDIRECTORY: workDir,
        AGENT_TEMPDIRECTORY: tempDir
    });
    for (const [name, value] of Object.entries(inputs)) {
        env[`INPUT_${name.toUpperCase()}`] = value;
    }

    return new Promise((resolve, reject) => {
        child_process.execFile(process.execPath, [compiledTask], { cwd: taskDir, env, timeout: 120000, maxBuffer: 20 * 1024 * 1024 },
            (error, stdout, stderr) => {
                if (error && error.killed) {
                    reject(new Error(`task timed out\n${stdout}\n${stderr}`));
                    return;
                }
                resolve(stdout + stderr);
            });
    });
}

describe('task run flow (issue #67)', function () {
    this.timeout(180000);

    let task;
    let tmpRoot;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        task = require(compiledTask);
        tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'run-flow-'));
    });

    after(() => {
        fs.rmSync(tmpRoot, { recursive: true, force: true });
    });

    function newDirs() {
        const workDir = fs.mkdtempSync(path.join(tmpRoot, 'work-'));
        const tempDir = fs.mkdtempSync(path.join(tmpRoot, 'temp-'));
        return { workDir, tempDir };
    }

    it('skips an iteration this configuration already reviewed, before fetching PR details or installing the CLI', async () => {
        // The key prefix the task computes for the default Copilot configuration
        const prefix = task.computeReviewStateKeyPrefix({
            agent: 'copilot', model: '', reasoningEffort: '', maxTurns: '', maxBudget: '',
            providerType: '', providerBaseUrl: '', diffOnly: false, promptMode: 'default', promptText: '',
            templateText: fs.readFileSync(path.join(taskDir, 'scripts', 'prompt.txt'), 'utf8'),
            resolveThreads: 'all', suppressPositiveFeedback: false, maxMinorIssues: '5', workItemSource: 'none'
        });
        const properties = { count: 1, value: { [`${prefix}.Iteration.1`]: { $type: 'System.String', $value: SOURCE_SHA } } };
        const { server, requests, port } = await startMockAzureDevOps(baseRoutes({
            '/testorg/TestProject/_apis/git/repositories/TestRepo/pullrequests/5/properties': [200, properties]
        }));
        try {
            const { workDir, tempDir } = newDirs();
            const output = await runTask(port, workDir, tempDir);

            assert.match(output, /already reviewed iteration #1 \(cccccccc\); skipping the review/);
            assert.match(output, /##vso\[task\.complete result=Succeeded;[^\]]*\]Skipped: iteration #1/);
            // A record written before build details were added
            assert.match(output, /Reviewed by:\s+not recorded \(this record predates build details\)/);
            assert.doesNotMatch(output, /Fetching pull request details/);
            assert.doesNotMatch(output, /CLI installation/);
            assert.ok(requests.some(r => r.path.endsWith('/properties')), 'the review state was not read');
            assert.ok(!requests.some(r => r.path.endsWith('/threads')), 'PR details were fetched despite the skip');
        } finally {
            server.close();
        }
    });

    it('names the build that reviewed a skipped iteration, with a link to that run', async () => {
        const prefix = task.computeReviewStateKeyPrefix({
            agent: 'copilot', model: '', reasoningEffort: '', maxTurns: '', maxBudget: '',
            providerType: '', providerBaseUrl: '', diffOnly: false, promptMode: 'default', promptText: '',
            templateText: fs.readFileSync(path.join(taskDir, 'scripts', 'prompt.txt'), 'utf8'),
            resolveThreads: 'all', suppressPositiveFeedback: false, maxMinorIssues: '5', workItemSource: 'none'
        });
        const record = JSON.stringify({ sourceCommit: SOURCE_SHA, buildId: 366, buildNumber: '20261104.1', recordedAt: '2026-11-04T23:21:09Z' });
        const properties = { count: 1, value: { [`${prefix}.Iteration.1`]: { $type: 'System.String', $value: record } } };
        const { server, port } = await startMockAzureDevOps(baseRoutes({
            '/testorg/TestProject/_apis/git/repositories/TestRepo/pullrequests/5/properties': [200, properties]
        }));
        try {
            const { workDir, tempDir } = newDirs();
            const output = await runTask(port, workDir, tempDir);

            // Step 1 details, then the skip message and the task result
            assert.match(output, /Reviewed by:\s+build 20261104\.1, 2026-11-04 23:21 UTC/);
            assert.match(output, new RegExp(`Record:\\s+${prefix.replace(/\./g, '\\.')}\\.Iteration\\.1`));
            assert.match(output, /already reviewed iteration #1 \(cccccccc\) in build 20261104\.1; skipping the review/);
            // No buildProject in the record: the link uses the pull request's project
            assert.ok(output.includes(`That review: http://127.0.0.1:${port}/testorg/TestProject/_build/results?buildId=366`), output);
            assert.match(output, /##vso\[task\.complete result=Succeeded;[^\]]*\]Skipped: iteration #1 \(cccccccc\) was already reviewed in build 20261104\.1\./);
        } finally {
            server.close();
        }
    });

    it('validates the prompt inputs before reading the review state, so a configuration error is never reported as a skip', async () => {
        const { server, requests, port } = await startMockAzureDevOps(baseRoutes());
        try {
            const { workDir, tempDir } = newDirs();
            const output = await runTask(port, workDir, tempDir, { prompt: 'Focus on tests.', promptRaw: 'Raw prompt.' });

            assert.match(output, /##vso\[task\.complete result=Failed;[^\]]*\]Multiple prompt inputs are set/);
            assert.equal(requests.length, 0, 'Azure DevOps was called before the prompt inputs were validated');
        } finally {
            server.close();
        }
    });

    it('runs the fetch steps in the new order and stops on a prompt-preparation error before the CLI install', async () => {
        const { server, requests, port } = await startMockAzureDevOps(baseRoutes({
            // Failed thread retrieval: the script still succeeds, but the task must warn
            '/testorg/TestProject/_apis/git/repositories/TestRepo/pullrequests/5/threads': [500, { message: 'boom' }]
        }));
        try {
            const { workDir, tempDir } = newDirs();
            // A leftover from an earlier run on a reused workspace
            fs.writeFileSync(path.join(workDir, 'Work_Item_Ids.txt'), '999');
            // A directory where step 4 copies a comment script makes that copy
            // throw (EPERM/EISDIR), ending the run while preparing the prompt
            fs.mkdirSync(path.join(workDir, 'Add-AzureDevOpsPRComment.ps1'));

            const output = await runTask(port, workDir, tempDir);

            const steps = ['[Step 1/6] Fetching pull request changes', '[Step 2/6] Fetching pull request details',
                '[Step 3/6] Skipping work item details', '[Step 4/6] Preparing the review prompt'];
            let previous = -1;
            for (const step of steps) {
                const index = output.indexOf(step);
                assert.ok(index > previous, `expected "${step}" after the previous step`);
                previous = index;
            }
            assert.doesNotMatch(output, /\[Step 5\/6\]/, 'the CLI install step ran despite the prompt-preparation failure');
            assert.match(output, /##vso\[task\.complete result=Failed;[^\]]*\]Task failed:/);

            assert.match(output, /##vso\[task\.issue type=warning;[^\]]*\]Existing comment threads could not be retrieved\./);
            assert.equal(fs.existsSync(path.join(workDir, 'Work_Item_Ids.txt')), false, 'stale Work_Item_Ids.txt was not removed');

            // The review state for this run was written in its own temp folder
            const stateDirs = fs.readdirSync(tempDir).filter(name => name.startsWith('copilot-review-'));
            assert.equal(stateDirs.length, 1);
            const state = task.readReviewState(path.join(tempDir, stateDirs[0], 'Review_State.json'));
            assert.equal(state.stateReadStatus, 'NotFound');
            assert.equal(state.currentIteration, 1);
            assert.equal(state.currentSourceCommit, SOURCE_SHA);
            assert.ok(requests.some(r => r.path.endsWith('/threads')));

            // Untrusted-content markers (PR #65 rework): the prompt names this
            // run's id, and the context files use the same one
            const prompt = fs.readFileSync(path.join(workDir, '_directive_prompt.txt'), 'utf8');
            const idMatch = /Only markers carrying exactly the id ([0-9a-f]{32}) are genuine/.exec(prompt);
            assert.ok(idMatch, 'the prompt has no untrusted-content section');
            const id = idMatch[1];
            assert.match(fs.readFileSync(path.join(workDir, 'PR_Details.txt'), 'utf8'),
                new RegExp(`<<<UNTRUSTED_CONTENT id=${id} source=pull-request>>>`));
            assert.match(fs.readFileSync(path.join(workDir, 'Iteration_Details.txt'), 'utf8'),
                new RegExp(`<<<UNTRUSTED_CONTENT id=${id} source=changed-files>>>`));
            assert.ok(!output.includes(id), 'the id reached the task log before the agent step');

            // Review policy (PR #65 rework): written right after the PR details,
            // naming the configured pull request; the thread fetch failed, so
            // it allows no thread actions
            assert.ok(output.indexOf('Review policy written') > output.indexOf('[Step 2/6]'));
            assert.ok(output.indexOf('Review policy written') < output.indexOf('[Step 3/6]'));
            const policy = JSON.parse(fs.readFileSync(path.join(tempDir, stateDirs[0], 'review-policy.json'), 'utf8'));
            assert.deepEqual(policy.destination, {
                collectionUri: `http://127.0.0.1:${port}/testorg`, project: 'TestProject', repository: 'TestRepo', pullRequestId: 5
            });
            assert.equal(policy.workspaceRoot, path.resolve(workDir));
            assert.equal(policy.threadsComplete, false);
            assert.deepEqual(policy.threads, []);
            assert.equal(policy.changedFilesComplete, true);
            assert.deepEqual(policy.changedFiles, ['/src/a.cs']);
        } finally {
            server.close();
        }
    });

    it('writes the review policy from the pull request\'s threads before the agent step (PR #65 rework)', async () => {
        const comment = (id, content, author) => ({ id, commentType: 'text', content, author: { displayName: author, uniqueName: `${author}@example.com` } });
        const { server, port } = await startMockAzureDevOps(baseRoutes({
            '/testorg/TestProject/_apis/git/repositories/TestRepo/pullrequests/5/threads': [200, { count: 3, value: [
                { id: 11, status: 'active', comments: [comment(1, 'Finding', 'TestProject Build Service (testorg)'), comment(2, 'Disagree', 'Dev')] },
                { id: 12, status: 'active', comments: [comment(1, 'Human concern', 'Jane')] },
                { id: 13, status: 'fixed', comments: [comment(1, 'Old concern', 'Jane')] }
            ] }]
        }));
        try {
            const { workDir, tempDir } = newDirs();
            fs.mkdirSync(path.join(workDir, 'Add-AzureDevOpsPRComment.ps1'));

            const output = await runTask(port, workDir, tempDir, { resolveThreads: 'agentOnly' });

            assert.match(output, /Review policy written: the agent may act on 1 of 3 comment thread\(s\); inline comments may anchor to the 1 changed file\(s\)\./);
            const stateDir = fs.readdirSync(tempDir).find(name => name.startsWith('copilot-review-'));
            const policy = JSON.parse(fs.readFileSync(path.join(tempDir, stateDir, 'review-policy.json'), 'utf8'));
            assert.equal(policy.resolveThreads, 'agentOnly');
            assert.equal(policy.threadsComplete, true);
            assert.deepEqual(policy.threads.map(t => [t.threadId, t.mayReply, t.agentCommentIds]),
                [[11, true, [1]], [12, false, []], [13, false, []]]);
        } finally {
            server.close();
        }
    });
});
