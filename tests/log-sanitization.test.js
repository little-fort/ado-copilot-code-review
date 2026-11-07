/**
 * Tests for pipeline logging-command neutralization in index.ts (issues #47, #69).
 *
 * The Azure Pipelines agent acts on `##vso[...]` found ANYWHERE in a stdout or
 * stderr line of a task. index.ts therefore routes every child process that
 * can print PR content (the PowerShell scripts and both review agents) through
 * forwardSanitizedOutput, a line-buffered filter in front of the pipeline log.
 *
 * Requires the compiled task (`npm run build`) — CI builds before testing.
 * Importing index.js is side-effect free: run() only fires when it is the
 * process entry point.
 */

const assert = require('node:assert/strict');
const child_process = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const { PassThrough, Writable } = require('node:stream');

const repoRoot = path.join(__dirname, '..');
const compiledTask = path.join(repoRoot, 'CopilotCodeReviewV1', 'index.js');

/** Writable that records every byte written to it. */
function collector() {
    const chunks = [];
    const stream = new Writable({
        write(chunk, _encoding, callback) {
            chunks.push(chunk);
            callback();
        }
    });
    stream.bytes = () => Buffer.concat(chunks);
    stream.text = () => Buffer.concat(chunks).toString('utf8');
    return stream;
}

describe('pipeline log sanitization (issue #69)', function () {
    let task;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        task = require(compiledTask);
    });

    describe('sanitizePipelineLog', () => {
        it('breaks a logging command in the middle of a line or after indentation', () => {
            assert.equal(
                task.sanitizePipelineLog("Write-Host '##vso[task.logissue type=error]boom'"),
                "Write-Host '# #vso[task.logissue type=error]boom'");
            assert.equal(
                task.sanitizePipelineLog('    ##vso[task.complete result=Failed;]'),
                '    # #vso[task.complete result=Failed;]');
        });

        it('breaks formatting markers anywhere', () => {
            assert.equal(task.sanitizePipelineLog('note ##[error]spoofed'), 'note # #[error]spoofed');
            assert.equal(task.sanitizePipelineLog('##[group]x'), '# #[group]x');
        });

        it('matches case-insensitively and breaks every occurrence', () => {
            const result = task.sanitizePipelineLog('a ##VSO[x]\nb ##vso[y] ##vso[z]');
            assert.equal(result, 'a # #VSO[x]\nb # #vso[y] # #vso[z]');
        });

        it('never leaves a marker behind in a run of hash characters', () => {
            assert.ok(!task.sanitizePipelineLog('###vso[a] ####vso[b]').includes('##vso['));
        });

        it('leaves ordinary text unchanged', () => {
            const text = '## Heading\nC# ## [link](url) #42';
            assert.equal(task.sanitizePipelineLog(text), text);
        });
    });

    describe('neutralizeLoggingCommands', () => {
        it('breaks ##vso[ anywhere but keeps the scripts\' own ##[error] diagnostics', () => {
            assert.equal(
                task.neutralizeLoggingCommands('##[error]API response: ##vso[task.complete result=Failed;]'),
                '##[error]API response: # #vso[task.complete result=Failed;]');
            assert.equal(task.neutralizeLoggingCommands('##[warning]note'), '##[warning]note');
        });
    });

    describe('forwardSanitizedStream', () => {
        it('sanitizes a marker split across two chunks', async () => {
            const source = new PassThrough();
            const target = collector();
            const done = task.forwardSanitizedStream(source, target, task.sanitizePipelineLog);

            source.write('prefix ##v');
            source.write('so[task.complete result=Failed;]x\nnext');
            source.end(' line');
            await done;

            assert.equal(target.text(), 'prefix # #vso[task.complete result=Failed;]x\nnext line');
        });

        it('forwards an overlong line in pieces instead of buffering it until the end', async () => {
            const source = new PassThrough();
            const target = collector();
            const done = task.forwardSanitizedStream(source, target, task.sanitizePipelineLog);

            // 80 KB with no newline exceeds the 64 KB hold-back limit
            source.write('a'.repeat(80 * 1024));
            await new Promise(resolve => setImmediate(resolve));
            assert.ok(target.bytes().length > 0, 'expected the overlong line to be forwarded before the stream ended');

            source.end('tail');
            await done;
            assert.equal(target.text(), 'a'.repeat(80 * 1024) + 'tail');
        });

        it('never splits a marker when forwarding an overlong line in pieces', async () => {
            // Each case ends the overlong line with a fragment that the next
            // chunk completes into a marker. A cut inside the marker would let
            // the agent see ##vso[ rejoined on one log line.
            for (const [ending, continuation] of [['##vs', 'o[task.complete]x'], ['a###', 'vso[x]'], ['#', '#vso[y]'], ['##', '[error]z']]) {
                const source = new PassThrough();
                const target = collector();
                const done = task.forwardSanitizedStream(source, target, task.sanitizePipelineLog);

                source.write('b'.repeat(70 * 1024) + ending);
                await new Promise(resolve => setImmediate(resolve));
                source.end(continuation);
                await done;

                const output = target.text();
                assert.ok(!output.includes('##vso[') && !output.includes('##['),
                    `marker survived for ending ${JSON.stringify(ending)}: ...${output.slice(-30)}`);
            }
        });

        it('pauses the source while a slow destination drains, without losing or reordering output', async () => {
            const source = new PassThrough();
            const chunks = [];
            // Small buffer and delayed acknowledgements force write() to report a full destination
            const slowTarget = new Writable({
                highWaterMark: 16,
                write(chunk, _encoding, callback) {
                    chunks.push(chunk);
                    setTimeout(callback, 2);
                }
            });
            let pauses = 0;
            const originalPause = source.pause.bind(source);
            source.pause = () => { pauses++; return originalPause(); };

            const done = task.forwardSanitizedStream(source, slowTarget, task.sanitizePipelineLog);
            const lines = Array.from({ length: 40 }, (_, i) => `line ${i} ##vso[task.setvariable variable=V${i}]x\n`);
            for (const line of lines) {
                source.write(line);
            }
            source.end();
            await done;
            // Let the destination work through its queued writes
            await new Promise(resolve => slowTarget.end(resolve));

            assert.ok(pauses > 0, 'expected the source to be paused at least once');
            assert.equal(Buffer.concat(chunks).toString('utf8'), task.sanitizePipelineLog(lines.join('')));
        });

        it('passes non-ASCII and non-UTF-8 bytes through unchanged', async () => {
            const source = new PassThrough();
            const target = collector();
            const done = task.forwardSanitizedStream(source, target, task.sanitizePipelineLog);

            // UTF-8 text plus raw bytes that are invalid UTF-8 (e.g. output in
            // a legacy console code page) must survive byte for byte
            const input = Buffer.concat([
                Buffer.from('café — 日本\n', 'utf8'),
                Buffer.from([0x80, 0xa9, 0xff, 0x0a])
            ]);
            source.end(input);
            await done;

            assert.deepEqual(target.bytes(), input);
        });
    });

    describe('forwardSanitizedOutput', function () {
        this.timeout(20000);

        it('sanitizes child stdout and stderr and resolves with the exit code', async () => {
            const child = child_process.spawn(process.execPath, ['-e', `
                process.stdout.write("out ##vso[task.setvariable variable=X]1\\n");
                process.stderr.write("err ##vso[task.complete result=Failed;]\\n");
                process.exitCode = 3;
            `], { stdio: ['ignore', 'pipe', 'pipe'] });

            const stdout = collector();
            const stderr = collector();
            const code = await task.forwardSanitizedOutput(child, task.sanitizePipelineLog, { stdout, stderr });

            assert.equal(code, 3);
            assert.equal(stdout.text(), 'out # #vso[task.setvariable variable=X]1\n');
            assert.equal(stderr.text(), 'err # #vso[task.complete result=Failed;]\n');
        });

        it('does not wait for a grandchild that keeps the pipes open', async () => {
            // The child starts a long-lived grandchild that inherits the
            // child's stdout/stderr, then exits — the situation a lingering
            // build server creates in a real review. On Linux/macOS the
            // inherited descriptors hold the pipes open, so only the grace
            // period lets the forwarder finish. On Windows, Node-created pipes
            // reach EOF when the child exits even with such a grandchild
            // (verified empirically), so there this is a smoke test.
            const child = child_process.spawn(process.execPath, ['-e', `
                const { spawn } = require('child_process');
                const grandchild = spawn(process.execPath, ['-e', 'setTimeout(() => {}, 60000)'], { stdio: 'inherit' });
                process.stdout.write('grandchild:' + grandchild.pid + '\\n', () => process.exit(0));
            `], { stdio: ['ignore', 'pipe', 'pipe'] });

            const stdout = collector();
            const started = Date.now();
            const code = await task.forwardSanitizedOutput(child, task.sanitizePipelineLog,
                { stdout, stderr: collector(), drainGraceMs: 300 });
            const elapsed = Date.now() - started;

            const match = stdout.text().match(/grandchild:(\d+)/);
            try {
                assert.equal(code, 0);
                assert.ok(match, `grandchild PID not reported; output: ${stdout.text()}`);
                assert.ok(elapsed < 10000, `resolved after ${elapsed}ms; expected the 300ms grace period to apply`);
                if (process.platform !== 'win32') {
                    // Proves the pipes really were held open, i.e. this test
                    // exercised the grace-period path rather than a normal EOF
                    assert.ok(elapsed >= 300, `resolved after ${elapsed}ms, before the grace period; the grandchild did not hold the pipes`);
                }
            } finally {
                if (match) {
                    try { process.kill(Number(match[1])); } catch { /* already exited */ }
                }
            }
        });
    });

    describe('output paths in index.ts', () => {
        const indexSource = fs.readFileSync(path.join(repoRoot, 'CopilotCodeReviewV1', 'index.ts'), 'utf8');

        it('no longer uses the start-of-line-only sanitizer anywhere', () => {
            const scriptsDir = path.join(repoRoot, 'CopilotCodeReviewV1', 'scripts');
            const sources = [['index.ts', indexSource]].concat(
                fs.readdirSync(scriptsDir)
                    .filter(name => name.endsWith('.ps1'))
                    .map(name => [name, fs.readFileSync(path.join(scriptsDir, name), 'utf8')]));

            for (const [name, source] of sources) {
                assert.ok(!source.includes("'(?m)^##'") && !source.includes('/^##/'),
                    `${name} still contains the start-of-line-only ## sanitizer (issue #69)`);
            }
        });

        it('inherits stdout only for the CLI installers', () => {
            // Any child spawned with stdio 'inherit' writes straight to the
            // pipeline log, bypassing forwardSanitizedOutput. Only the
            // installers (which print no PR content) may do that.
            const inheritCount = (indexSource.match(/stdio:\s*'inherit'/g) || []).length;
            assert.equal(inheritCount, 1,
                'expected exactly one stdio: \'inherit\' spawn (runInstallerOnce); route new child processes through forwardSanitizedOutput');
            assert.match(indexSource, /function runInstallerOnce[\s\S]*?stdio:\s*'inherit'/);
        });
    });
});
