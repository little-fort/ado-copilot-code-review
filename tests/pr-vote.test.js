const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const task = require('../CopilotCodeReviewV1/index.js');
const snapshot = {
    reviewerId: '11111111-2222-3333-4444-555555555555',
    previousVote: -10,
    sourceCommit: 'a'.repeat(40),
    targetCommit: 'b'.repeat(40)
};

describe('automatic PR voting', () => {
    let tempDir;
    before(() => { tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pr-vote-')); });
    after(() => { fs.rmSync(tempDir, { recursive: true, force: true }); });

    function writeResult(overrides = {}) {
        const file = path.join(tempDir, 'result.json');
        fs.writeFileSync(file, JSON.stringify({
            runId: 'current-run', sourceCommit: snapshot.sourceCommit, targetCommit: snapshot.targetCommit,
            complete: true, previousFindingsReviewed: true, highestSeverity: 'major',
            summary: 'The Critical issue is fixed; a Major issue remains.',
            ...overrides
        }));
        return file;
    }

    it('allows votes to improve or worsen based on the current overall severity', () => {
        for (const previousVote of [-10, -5, 0, 5, 10]) {
            for (const [highestSeverity, vote] of [['none', 10], ['minor', 5], ['major', -5], ['critical', -10]]) {
                const result = task.readVoteResult(writeResult({ highestSeverity }), { ...snapshot, previousVote }, 'current-run');
                assert.equal(result.vote, vote, `${previousVote} -> ${highestSeverity}`);
            }
        }
    });

    it('refuses incomplete results and results from a different run or PR snapshot', () => {
        for (const overrides of [
            { complete: false }, { previousFindingsReviewed: false }, { runId: 'old-run' },
            { sourceCommit: 'c'.repeat(40) }, { targetCommit: 'd'.repeat(40) },
            { highestSeverity: 'unknown' }, { summary: '' }
        ]) {
            assert.throws(() => task.readVoteResult(writeResult(overrides), snapshot, 'current-run'), /PR vote unchanged/);
        }
        assert.throws(() => task.readVoteResult(path.join(tempDir, 'missing.json'), snapshot, 'current-run'));
    });

    it('requests a full overall assessment without treating missing new comments as approval', () => {
        const directive = task.buildVoteDirective(snapshot, 'current-run', '_vote.json');
        assert.match(directive, /previous vote is Rejected \(-10\)/);
        assert.match(directive, /ALL existing unresolved review findings/);
        assert.match(directive, /outside the incremental diff/);
        assert.match(directive, /Missing new comments NEVER means approval/);
        assert.match(directive, /Minor concerns even if the comment cap suppresses them/);
        assert.match(directive, /Do NOT cast a vote yourself/);
    });
});
