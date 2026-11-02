/**
 * Tests for the CLI launch hardening in index.ts (PR #65 rework, Part A):
 * the flags and environment each CLI gets, the saved-trust check for Copilot,
 * and the version/update helpers behind autoUpdateCli.
 *
 * Requires the compiled task (`npm run build`). The flag builders return a
 * fragment of the PowerShell command that launches the CLI, so besides string
 * checks, the "as PowerShell parses it" tests run the fragment through a real
 * pwsh and inspect the argv a program would receive.
 */

const assert = require('node:assert/strict');
const child_process = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

describe('CLI launch hardening (PR #65 rework)', function () {
    this.timeout(30000);

    let task;
    let tmpDir;
    let argvEcho;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        task = require(compiledTask);
        tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cli-hardening-'));
        argvEcho = path.join(tmpDir, 'argv-echo.js');
        fs.writeFileSync(argvEcho, 'process.stdout.write(JSON.stringify(process.argv.slice(2)));');
    });

    after(() => {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    /** The argv a program receives when `flags` follows it in a pwsh -Command line. */
    function argvAsPowerShellPassesIt(flags) {
        const node = process.execPath.replace(/'/g, "''");
        const script = argvEcho.replace(/'/g, "''");
        const result = child_process.spawnSync('pwsh', ['-NoProfile', '-Command', `& '${node}' '${script}' ${flags}`],
            { encoding: 'utf8' });
        assert.equal(result.status, 0, `pwsh failed: ${result.stderr}`);
        return JSON.parse(result.stdout);
    }

    describe('buildClaudeCliArgs', () => {
        const base = {
            diffOnlyActive: false,
            ignoreCustomInstructions: false,
            workingDirectory: path.join(os.tmpdir(), "agent work", "it's a repo")
        };

        it('always loads only the user setting source', () => {
            for (const ignoreCustomInstructions of [false, true]) {
                for (const diffOnlyActive of [false, true]) {
                    const flags = task.buildClaudeCliArgs({ ...base, ignoreCustomInstructions, diffOnlyActive });
                    assert.match(flags, /--setting-sources user(\s|$)/);
                    assert.doesNotMatch(flags, /--safe-mode|--strict-mcp-config|--bare|--restricted/);
                }
            }
        });

        it('adds the checkout as an extra directory only when repository instructions are kept', () => {
            const keep = argvAsPowerShellPassesIt(task.buildClaudeCliArgs(base));
            const addDir = keep.indexOf('--add-dir');
            assert.ok(addDir >= 0, 'expected --add-dir when instructions are kept');
            // Spaces and the apostrophe survive PowerShell parsing as one argument
            assert.equal(keep[addDir + 1], base.workingDirectory);

            const ignore = argvAsPowerShellPassesIt(task.buildClaudeCliArgs({ ...base, ignoreCustomInstructions: true }));
            assert.ok(!ignore.includes('--add-dir'), 'no --add-dir when instructions are ignored');
        });

        it('keeps the existing tool rules and options', () => {
            const full = argvAsPowerShellPassesIt(task.buildClaudeCliArgs({
                ...base, model: 'sonnet', reasoningEffort: 'high', maxTurns: '20', maxBudget: '2'
            }));
            assert.deepEqual(full.slice(0, 5),
                ['--dangerously-skip-permissions', '--output-format', 'stream-json', '--verbose', '--include-partial-messages']);
            const allowed = full.indexOf('--allowedTools');
            assert.deepEqual(full.slice(allowed + 1, allowed + 7), ['Bash', 'Read', 'Write', 'Edit', 'Glob', 'Grep']);
            assert.equal(full[full.indexOf('--disallowedTools') + 1], 'Bash(git push *)');
            for (const [flag, value] of [['--model', 'sonnet'], ['--effort', 'high'], ['--max-turns', '20'], ['--max-budget-usd', '2']]) {
                assert.equal(full[full.indexOf(flag) + 1], value, `${flag} value`);
            }

            const diffOnly = argvAsPowerShellPassesIt(task.buildClaudeCliArgs({ ...base, diffOnlyActive: true }));
            const diffAllowed = diffOnly.indexOf('--allowedTools');
            assert.deepEqual(diffOnly.slice(diffAllowed + 1, diffAllowed + 3), ['Bash(pwsh *)', 'Write']);
        });
    });

    describe('buildClaudeCliEnv', () => {
        it('turns on CLAUDE.md loading for the extra directory only when instructions are kept', () => {
            const baseEnv = { PATH: '/usr/bin', ANTHROPIC_API_KEY: 'key' };
            const keep = task.buildClaudeCliEnv(baseEnv, { ignoreCustomInstructions: false });
            assert.equal(keep.CLAUDE_CODE_ADDITIONAL_DIRECTORIES_CLAUDE_MD, '1');
            assert.equal(keep.ANTHROPIC_API_KEY, 'key');

            const ignore = task.buildClaudeCliEnv(baseEnv, { ignoreCustomInstructions: true });
            assert.equal(ignore.CLAUDE_CODE_ADDITIONAL_DIRECTORIES_CLAUDE_MD, undefined);
            assert.equal(baseEnv.CLAUDE_CODE_ADDITIONAL_DIRECTORIES_CLAUDE_MD, undefined, 'base env must not be mutated');
        });
    });

    describe('buildCopilotCliArgs', () => {
        it('disables built-in MCP servers and strips the CLI credentials from the agent shell', () => {
            const argv = argvAsPowerShellPassesIt(task.buildCopilotCliArgs({ ignoreCustomInstructions: false }));
            assert.deepEqual(argv.slice(0, 4), ['--allow-all-paths', '--allow-all-tools', '--deny-tool', 'shell(git push)']);
            assert.ok(argv.includes('--disable-builtin-mcps'));
            const secret = argv.filter(arg => arg.startsWith('--secret-env-vars='));
            assert.equal(secret.length, 1, 'one --secret-env-vars argument');
            const names = secret[0].slice('--secret-env-vars='.length).split(',');
            assert.deepEqual(names, task.COPILOT_SECRET_ENV_VARS);
            assert.deepEqual(names, ['GH_TOKEN', 'COPILOT_GITHUB_TOKEN', 'GITHUB_TOKEN', 'COPILOT_PROVIDER_API_KEY', 'COPILOT_PROVIDER_BEARER_TOKEN']);
            // The comment scripts the agent runs read it from the shell environment
            assert.ok(!names.includes('AZUREDEVOPS_TOKEN'));
            // No effect alongside --allow-all-paths, so it is deliberately not passed
            assert.ok(!argv.includes('--disallow-temp-dir'));
            assert.ok(!argv.includes('--no-custom-instructions'));
            assert.ok(!argv.includes('--excluded-tools'));
        });

        it('drops repository instructions and skills when asked', () => {
            const argv = argvAsPowerShellPassesIt(task.buildCopilotCliArgs({ ignoreCustomInstructions: true }));
            assert.ok(argv.includes('--no-custom-instructions'));
            assert.equal(argv[argv.indexOf('--excluded-tools') + 1], 'skill');
        });

        it('keeps the model and reasoning effort options', () => {
            const argv = argvAsPowerShellPassesIt(task.buildCopilotCliArgs({
                ignoreCustomInstructions: false, model: 'gpt-5.5', reasoningEffort: 'xhigh'
            }));
            assert.equal(argv[argv.indexOf('--model') + 1], 'gpt-5.5');
            assert.equal(argv[argv.indexOf('--reasoning-effort') + 1], 'xhigh');
        });
    });

    describe('buildCopilotCliEnv', () => {
        it('removes the variables that make the CLI trust the checkout and keeps the rest', () => {
            const baseEnv = {
                PATH: '/usr/bin',
                GH_TOKEN: 'token',
                COPILOT_MODEL: 'gpt-5.5',
                copilot_allow_all: 'true',
                GITHUB_COPILOT_PROMPT_MODE_REPO_HOOKS: 'true',
                GITHUB_COPILOT_PROMPT_MODE_WORKSPACE_MCP: 'true',
                GITHUB_COPILOT_PROMPT_MODE_EXTENSIONS: 'true'
            };
            const { env, removed } = task.buildCopilotCliEnv(baseEnv);
            assert.deepEqual(removed.sort(), [
                'GITHUB_COPILOT_PROMPT_MODE_EXTENSIONS',
                'GITHUB_COPILOT_PROMPT_MODE_REPO_HOOKS',
                'GITHUB_COPILOT_PROMPT_MODE_WORKSPACE_MCP',
                'copilot_allow_all'
            ]);
            assert.deepEqual(env, { PATH: '/usr/bin', GH_TOKEN: 'token', COPILOT_MODEL: 'gpt-5.5' });
            assert.equal(baseEnv.copilot_allow_all, 'true', 'base env must not be mutated');
        });
    });

    describe('findTrustingFolder', () => {
        it('matches the folder itself and anything below it on Windows, ignoring case', () => {
            const checkout = 'D:\\a\\1\\s';
            assert.equal(task.findTrustingFolder(checkout, ['D:\\a\\1\\s'], 'win32'), 'D:\\a\\1\\s');
            assert.equal(task.findTrustingFolder(checkout, ['d:\\A'], 'win32'), 'd:\\A');
            assert.equal(task.findTrustingFolder(checkout, ['D:/a/'], 'win32'), 'D:/a/');
            assert.equal(task.findTrustingFolder(checkout, ['D:\\'], 'win32'), 'D:\\');
        });

        it('does not treat a sibling with a shared prefix as a parent', () => {
            assert.equal(task.findTrustingFolder('C:\\work\\ab', ['C:\\work\\a'], 'win32'), null);
            assert.equal(task.findTrustingFolder('/home/agent/work-2', ['/home/agent/work'], 'linux'), null);
        });

        it('compares case-sensitively on Linux', () => {
            assert.equal(task.findTrustingFolder('/home/agent/_work/1/s', ['/home/Agent'], 'linux'), null);
            assert.equal(task.findTrustingFolder('/home/agent/_work/1/s', ['/home/agent'], 'linux'), '/home/agent');
            assert.equal(task.findTrustingFolder('/home/agent/_work/1/s', ['/'], 'linux'), '/');
        });

        it('ignores a missing or malformed trustedFolders value', () => {
            assert.equal(task.findTrustingFolder('C:\\x', undefined, 'win32'), null);
            assert.equal(task.findTrustingFolder('C:\\x', 'C:\\', 'win32'), null);
            assert.equal(task.findTrustingFolder('C:\\x', [42, null, '', '  '], 'win32'), null);
        });
    });

    describe('reading the Copilot config', () => {
        it('parses JSON with comments and trailing commas, leaving strings intact', () => {
            const text = [
                '// User settings belong in settings.json.',
                '// This file is managed automatically.',
                '{',
                '  /* block',
                '     comment */',
                '  "trustedFolders": ["C:\\\\work // not a comment", "/srv/a,b", "/* nor this */",],',
                '  "quote": "say \\"hi\\" // still a string",',
                '}'
            ].join('\n');
            assert.deepEqual(task.parseJsonWithComments(text), {
                trustedFolders: ['C:\\work // not a comment', '/srv/a,b', '/* nor this */'],
                quote: 'say "hi" // still a string'
            });
        });

        it('reads trustedFolders from COPILOT_HOME, and tolerates a missing or broken config', () => {
            const home = path.join(tmpDir, 'copilot-home');
            fs.mkdirSync(home, { recursive: true });
            const saved = process.env.COPILOT_HOME;
            process.env.COPILOT_HOME = home;
            try {
                assert.equal(task.readCopilotTrustedFolders(), undefined, 'no config.json');
                fs.writeFileSync(path.join(home, 'config.json'),
                    '\uFEFF// This file is managed automatically.\n{ "trustedFolders": ["D:\\\\a"], "firstLaunchAt": "x" }');
                assert.deepEqual(task.readCopilotTrustedFolders(), ['D:\\a']);
                fs.writeFileSync(path.join(home, 'config.json'), '{ not json');
                assert.equal(task.readCopilotTrustedFolders(), undefined, 'broken config');
            } finally {
                if (saved === undefined) {
                    delete process.env.COPILOT_HOME;
                } else {
                    process.env.COPILOT_HOME = saved;
                }
            }
        });
    });

    describe('CLI versions', () => {
        it('parses the --version output of both CLIs', () => {
            assert.equal(task.parseCliVersion("GitHub Copilot CLI 1.0.105.\nRun 'copilot update' to check for updates."), '1.0.105');
            assert.equal(task.parseCliVersion('2.1.316 (Claude Code)'), '2.1.316');
            assert.equal(task.parseCliVersion('command not found'), null);
            assert.equal(task.parseCliVersion(null), null);
        });

        it('compares versions numerically', () => {
            assert.ok(task.compareVersions('1.0.9', '1.0.41') < 0);
            assert.ok(task.compareVersions('2.10.0', '2.9.9') > 0);
            assert.equal(task.compareVersions('2.1.69', '2.1.69'), 0);
        });

        it('warns only for a known version below the floor', () => {
            assert.equal(task.COPILOT_CLI_FLOOR.minimum, '1.0.41');
            assert.equal(task.CLAUDE_CLI_FLOOR.minimum, '2.1.69');
            const warning = task.cliVersionWarning('GitHub Copilot CLI', '1.0.39', task.COPILOT_CLI_FLOOR);
            assert.match(warning, /1\.0\.39 is older than 1\.0\.41/);
            assert.match(warning, /autoUpdateCli: true/);
            assert.equal(task.cliVersionWarning('GitHub Copilot CLI', '1.0.41', task.COPILOT_CLI_FLOOR), null);
            assert.equal(task.cliVersionWarning('Claude Code CLI', null, task.CLAUDE_CLI_FLOOR), null);
        });
    });

    describe('updateExistingCli (autoUpdateCli)', () => {
        function stub(name, body) {
            const file = path.join(tmpDir, name);
            fs.writeFileSync(file, body);
            return file;
        }

        it('does not claim the CLI is current when the version did not change', async () => {
            // Claude has exited 0 without updating, so "unchanged" must stay ambiguous
            const result = await task.updateExistingCli({
                label: 'Claude Code CLI',
                getVersion: () => '2.1.316',
                command: process.execPath,
                args: [stub('noop-update.js', 'process.exit(0);')],
                maxAttempts: 1
            });
            assert.equal(result.ok, true);
            assert.equal(result.message,
                'Claude Code CLI version unchanged (2.1.316); it may already be current, or the update did not apply.');
        });

        it('reports the new version after an update', async () => {
            const versions = ['1.0.100', '1.0.105'];
            const result = await task.updateExistingCli({
                label: 'GitHub Copilot CLI',
                getVersion: () => versions.shift(),
                command: process.execPath,
                args: [stub('ok-update.js', 'process.exit(0);')],
                maxAttempts: 1
            });
            assert.deepEqual(result, { ok: true, message: 'GitHub Copilot CLI updated from 1.0.100 to 1.0.105.' });
        });

        it('keeps the installed version when the update fails or hangs', async () => {
            const failed = await task.updateExistingCli({
                label: 'GitHub Copilot CLI',
                getVersion: () => '1.0.100',
                command: process.execPath,
                args: [stub('failing-update.js', 'process.exit(1);')],
                maxAttempts: 1
            });
            assert.equal(failed.ok, false);
            assert.match(failed.message, /^Could not update the GitHub Copilot CLI: .*Exit code: 1/);
            assert.match(failed.message, /Continuing with the installed version \(1\.0\.100\)\.$/);

            const hung = await task.updateExistingCli({
                label: 'Claude Code CLI',
                getVersion: () => '2.1.316',
                command: process.execPath,
                args: [stub('hanging-update.js', 'setTimeout(() => {}, 20000);')],
                maxAttempts: 1,
                timeoutMs: 500
            });
            assert.equal(hung.ok, false);
            assert.match(hung.message, /Timed out after 1 seconds/);
        });

        it('stops the whole update process tree on a timeout, not just the pwsh wrapper', async () => {
            // Same launch shape as the task: pwsh -Command '<cli> update', where
            // the CLI (here a node stub) is a grandchild of the task process
            const pidFile = path.join(tmpDir, 'grandchild.pid');
            const grandchild = stub('hanging-grandchild.js',
                `require('fs').writeFileSync(${JSON.stringify(pidFile)}, String(process.pid)); setTimeout(() => {}, 60000);`);
            const quote = value => `'${value.replace(/'/g, "''")}'`;
            const result = await task.updateExistingCli({
                label: 'GitHub Copilot CLI',
                getVersion: () => '1.0.100',
                command: 'pwsh',
                args: ['-NoProfile', '-Command', `& ${quote(process.execPath)} ${quote(grandchild)}; exit $LASTEXITCODE`],
                maxAttempts: 1,
                timeoutMs: 4000
            });
            assert.equal(result.ok, false);
            assert.match(result.message, /Timed out/);

            const pid = Number(fs.readFileSync(pidFile, 'utf8'));
            const isAlive = () => { try { process.kill(pid, 0); return true; } catch { return false; } };
            for (let i = 0; i < 20 && isAlive(); i++) {
                await new Promise(resolve => setTimeout(resolve, 100));
            }
            assert.equal(isAlive(), false, 'the grandchild updater must not survive the timeout');
        });

        it('describes an update whose version could not be read', () => {
            assert.equal(task.describeCliUpdate('Claude Code CLI', null, '2.1.316'),
                'Claude Code CLI update finished, but the version could not be read (now 2.1.316).');
            assert.equal(task.describeCliUpdate('Claude Code CLI', '2.1.300', null),
                'Claude Code CLI update finished, but the version could not be read.');
        });
    });
});
