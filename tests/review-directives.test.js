/**
 * Tests for the review behavior directives builder in index.ts
 * (issues #55 — thread resolution scope, and #28 — positive feedback
 * suppression).
 *
 * Requires the compiled task (`npm run build`) — CI builds before testing.
 */

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

describe('buildReviewDirectives (issues #55, #28)', () => {
    let buildReviewDirectives;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        ({ buildReviewDirectives } = require(compiledTask));
    });

    it('always emits the directives section header and precedence note', () => {
        const section = buildReviewDirectives('agentOnly', false);
        assert.match(section, /# Review Behavior Directives/);
        assert.match(section, /take precedence over any conflicting instructions/);
    });

    it('agentOnly forbids touching human reviewers\' threads', () => {
        const section = buildReviewDirectives('agentOnly', false);
        assert.match(section, /NEVER resolve, close, or otherwise change the status of threads created by human reviewers/);
        // Scopes allowed resolution to the agent-thread JSON section
        assert.match(section, /COPILOT COMMENT THREADS \(JSON\)/);
    });

    it('all permits resolving addressed human threads with the two-step process', () => {
        const section = buildReviewDirectives('all', false);
        assert.match(section, /may resolve threads created by human reviewers/);
        assert.match(section, /first reply to the thread/);
        assert.doesNotMatch(section, /NEVER resolve/);
    });

    it('includes the positive feedback directive only when suppression is enabled', () => {
        const withSuppression = buildReviewDirectives('agentOnly', true);
        assert.match(withSuppression, /Do NOT post positive, congratulatory/);
        assert.match(withSuppression, /overrides any earlier instruction to leave a comment indicating the code looks good/);

        const withoutSuppression = buildReviewDirectives('agentOnly', false);
        assert.doesNotMatch(withoutSuppression, /Positive feedback suppression/);
    });

    it('never contains double quotes, which are disallowed in prompt content', () => {
        // Both directive combinations must stay quote-free — the prompt pipeline
        // has historically had double-quote escaping constraints
        for (const [scope, suppress] of [['agentOnly', true], ['all', true]]) {
            assert.doesNotMatch(buildReviewDirectives(scope, suppress), /"/);
        }
    });
});
