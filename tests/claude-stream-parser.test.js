/**
 * Tests for the Claude Code stream parser in index.ts (buildClaudeStreamParser).
 *
 * Claude Code's stream-json output delivers review text as many small
 * text_delta fragments. The parser previously printed each fragment with
 * Write-Host's implicit newline, chopping the review into one fragment per
 * log line. These tests run the generated PowerShell against a recorded
 * event stream (replayed with Get-Content in place of the real CLI) and check
 * the exact log text.
 *
 * Requires the compiled task (`npm run build`) and pwsh on PATH.
 */

const assert = require('node:assert/strict');
const child_process = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

/** Wraps a Claude Code stream event in the stream-json envelope. */
function streamEvent(event) {
    return JSON.stringify({ type: 'stream_event', event });
}

function textDelta(index, text) {
    return streamEvent({ type: 'content_block_delta', index, delta: { type: 'text_delta', text } });
}

describe('buildClaudeStreamParser', function () {
    this.timeout(30000);

    let buildClaudeStreamParser;
    let tmpDir;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        ({ buildClaudeStreamParser } = require(compiledTask));
        tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'claude-stream-'));
    });

    after(() => {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    /** Replays the given stream-json lines through the parser; returns its stdout. */
    function runParser(lines, env = {}) {
        const fixture = path.join(tmpDir, `stream-${Date.now()}-${Math.random().toString(16).slice(2)}.jsonl`);
        fs.writeFileSync(fixture, lines.join('\n') + '\n', 'utf8');

        const script = buildClaudeStreamParser(`Get-Content -LiteralPath '${fixture}'`);
        const result = child_process.spawnSync('pwsh', ['-NoProfile', '-Command', script],
            { encoding: 'utf8', env: { ...process.env, ...env } });
        assert.equal(result.status, 0, `parser exited ${result.status}: ${result.stderr}`);
        // Write-Host emits CRLF on Windows; compare platform-independently
        return result.stdout.replace(/\r\n/g, '\n');
    }

    it('joins streamed text fragments into readable lines and separates tool calls', () => {
        const output = runParser([
            JSON.stringify({ type: 'system', subtype: 'init' }),
            streamEvent({ type: 'content_block_start', index: 0, content_block: { type: 'text', text: '' } }),
            textDelta(0, 'Reviewing the'),
            textDelta(0, ' changes now.\nFirst '),
            textDelta(0, 'finding: missing null check.'),
            streamEvent({ type: 'content_block_stop', index: 0 }),
            streamEvent({ type: 'content_block_start', index: 1, content_block: { type: 'tool_use', name: 'Bash' } }),
            streamEvent({ type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: '{"command":' } }),
            streamEvent({ type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: '"git diff"}' } }),
            streamEvent({ type: 'content_block_stop', index: 1 }),
            streamEvent({ type: 'content_block_start', index: 2, content_block: { type: 'text', text: '' } }),
            textDelta(2, 'Summary line\n'),
            streamEvent({ type: 'content_block_stop', index: 2 }),
            JSON.stringify({ type: 'result', subtype: 'success' })
        ]);

        assert.equal(output,
            'Reviewing the changes now.\n' +
            'First finding: missing null check.\n' +
            '\n' +
            '    [Bash] {"command":"git diff"}\n' +
            // A block whose text already ends with a newline gets no extra one
            'Summary line\n' +
            '\n');
    });

    it('shows a non-JSON line verbatim on its own line', () => {
        const output = runParser([
            streamEvent({ type: 'content_block_start', index: 0, content_block: { type: 'text', text: '' } }),
            textDelta(0, 'Partial'),
            'Error: plain-text message from the CLI',
            textDelta(0, ' done.'),
            streamEvent({ type: 'content_block_stop', index: 0 })
        ]);

        assert.equal(output,
            'Partial\n' +
            'Error: plain-text message from the CLI\n' +
            ' done.\n' +
            '\n');
    });

    describe('error results (issue #67)', () => {
        // The final result event says whether the review stopped early (e.g. at
        // --max-turns or --max-budget-usd). Claude Code 2.1.316 also exits 1 in
        // those cases (verified), but the parser leaves a marker that blocks
        // recording the run regardless of the exit code, as defense in depth.
        function markerPath(stateDir) {
            return path.join(stateDir, 'activity', 'claude-result.agent-error');
        }

        it('leaves an agent-error marker for an error result without changing the output', () => {
            const stateDir = fs.mkdtempSync(path.join(tmpDir, 'state-'));
            const output = runParser([
                textDelta(0, 'Partial review'),
                JSON.stringify({ type: 'result', subtype: 'error_max_turns', is_error: false })
            ], { COPILOT_REVIEW_STATE_DIR: stateDir });

            assert.ok(fs.existsSync(markerPath(stateDir)), 'expected an agent-error marker');
            assert.equal(fs.readFileSync(markerPath(stateDir), 'utf8'), 'error_max_turns');
            assert.equal(output, 'Partial review\n\n');
        });

        it('leaves no marker for a successful result', () => {
            const stateDir = fs.mkdtempSync(path.join(tmpDir, 'state-'));
            runParser([JSON.stringify({ type: 'result', subtype: 'success', is_error: false })], { COPILOT_REVIEW_STATE_DIR: stateDir });
            assert.equal(fs.existsSync(markerPath(stateDir)), false);
        });
    });
});
