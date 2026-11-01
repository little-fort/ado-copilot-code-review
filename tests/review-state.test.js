/**
 * Tests for the incremental re-review decisions in index.ts (issue #67).
 *
 * The task records a reviewed iteration as a pull request property, skips an
 * iteration that was already reviewed, and points re-reviews at what changed.
 * Recording wrongly would let later runs skip or narrow past code nobody
 * reviewed, so the recording rule is tested condition by condition.
 *
 * Requires the compiled task (`npm run build`); importing index.js is side-effect free.
 */

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

const SHA_A = 'a'.repeat(40);
const SHA_C = 'c'.repeat(40);

describe('incremental re-review state (issue #67)', () => {
    let task;
    let tmpDir;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        task = require(compiledTask);
        tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'review-state-'));
    });

    after(() => {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    /** A fresh folder shaped like the task's per-run state folder. */
    function newStateDir() {
        return fs.mkdtempSync(path.join(tmpDir, 'run-'));
    }

    function baseConfiguration() {
        return {
            agent: 'copilot',
            model: '',
            reasoningEffort: '',
            maxTurns: '',
            maxBudget: '',
            providerType: '',
            providerBaseUrl: '',
            diffOnly: false,
            promptMode: 'default',
            promptText: '',
            templateText: 'Review this pull request.\n',
            resolveThreads: 'all',
            suppressPositiveFeedback: false,
            maxMinorIssues: '5',
            workItemSource: 'none'
        };
    }

    describe('computeReviewStateKeyPrefix', () => {
        it('produces a stable prefix in the property-name format', () => {
            const prefix = task.computeReviewStateKeyPrefix(baseConfiguration());
            assert.match(prefix, /^CopilotCodeReview\.v1\.[0-9a-f]{16}$/);
            assert.equal(task.computeReviewStateKeyPrefix(baseConfiguration()), prefix);
        });

        it('changes with every review-affecting setting', () => {
            const base = task.computeReviewStateKeyPrefix(baseConfiguration());
            const changes = {
                agent: 'claude',
                model: 'gpt-5.5',
                reasoningEffort: 'high',
                maxTurns: '20',
                maxBudget: '2',
                providerType: 'openai',
                providerBaseUrl: 'https://llm.example.com/v1',
                diffOnly: true,
                promptMode: 'prompt',
                promptText: 'Focus on security.',
                templateText: 'A different template.',
                resolveThreads: 'agentOnly',
                suppressPositiveFeedback: true,
                maxMinorIssues: '0',
                workItemSource: 'azureBoards'
            };
            for (const [setting, value] of Object.entries(changes)) {
                const changed = task.computeReviewStateKeyPrefix({ ...baseConfiguration(), [setting]: value });
                assert.notEqual(changed, base, `changing ${setting} must change the key prefix`);
            }
        });

        it('ignores line-ending and surrounding-whitespace differences in prompt text', () => {
            const base = task.computeReviewStateKeyPrefix({ ...baseConfiguration(), promptText: 'Line one\nLine two' });
            const crlf = task.computeReviewStateKeyPrefix({ ...baseConfiguration(), promptText: '  Line one\r\nLine two\r\n' });
            assert.equal(crlf, base);
        });
    });

    describe('readReviewState', () => {
        function writeState(contents) {
            const file = path.join(newStateDir(), 'Review_State.json');
            fs.writeFileSync(file, contents, 'utf8');
            return file;
        }

        it('returns null for a missing or malformed file', () => {
            assert.equal(task.readReviewState(path.join(tmpDir, 'missing.json')), null);
            assert.equal(task.readReviewState(writeState('not json')), null);
            assert.equal(task.readReviewState(writeState('{"currentIteration":"3"}')), null);
        });

        it('reads a file with a byte-order mark and drops unusable delta entries', () => {
            const state = task.readReviewState(writeState('﻿' + JSON.stringify({
                currentIteration: 3,
                currentSourceCommit: SHA_C,
                fullChangesStatus: 'Complete',
                stateReadStatus: 'Found',
                lastReviewedIteration: 2,
                lastReviewedSourceCommit: SHA_A,
                delta: { status: 'Complete', reason: null, files: [{ path: '/src/a.cs', changeType: 'edit' }, { changeType: 'edit' }, null] }
            })));
            assert.equal(state.currentIteration, 3);
            assert.equal(state.lastReviewedIteration, 2);
            assert.deepEqual(state.delta.files, [{ path: '/src/a.cs', changeType: 'edit' }]);
        });
    });

    describe('isIterationAlreadyReviewed', () => {
        const reviewed = {
            currentIteration: 3, currentSourceCommit: SHA_C, fullChangesStatus: 'Complete',
            stateReadStatus: 'Found', lastReviewedIteration: 3, lastReviewedSourceCommit: SHA_C.toUpperCase(),
            delta: { status: 'NotApplicable', reason: null, files: [] }
        };

        it('is true only for the same iteration and commit', () => {
            assert.equal(task.isIterationAlreadyReviewed(reviewed), true);
            assert.equal(task.isIterationAlreadyReviewed({ ...reviewed, lastReviewedIteration: 2 }), false);
            assert.equal(task.isIterationAlreadyReviewed({ ...reviewed, lastReviewedSourceCommit: SHA_A }), false);
            assert.equal(task.isIterationAlreadyReviewed({ ...reviewed, stateReadStatus: 'NotFound' }), false);
            assert.equal(task.isIterationAlreadyReviewed({ ...reviewed, currentSourceCommit: '' }), false);
            assert.equal(task.isIterationAlreadyReviewed(null), false);
        });
    });

    describe('summarizeReviewActivity', () => {
        function writeActivity(stateDir, files) {
            const activityDir = path.join(stateDir, 'activity');
            fs.mkdirSync(activityDir, { recursive: true });
            for (const [name, contents] of Object.entries(files)) {
                fs.writeFileSync(path.join(activityDir, name), contents);
            }
        }
        const id = n => String(n).repeat(32).slice(0, 32).replace(/[^0-9a-f]/g, 'a');

        it('reports no activity when nothing was posted', () => {
            assert.deepEqual(task.summarizeReviewActivity(newStateDir()),
                { readable: true, completedPosts: 0, unfinishedAttempts: 0, agentErrors: 0 });
        });

        it('counts completed posts, but not the wrapper that delegates them', () => {
            const stateDir = newStateDir();
            writeActivity(stateDir, {
                [`${id(1)}.started`]: 'thread', [`${id(1)}.done`]: 'thread',
                [`${id(2)}.started`]: 'reply', [`${id(2)}.done`]: 'reply',
                [`${id(3)}.started`]: 'update', [`${id(3)}.done`]: 'update',
                [`${id(4)}.started`]: 'wrapper', [`${id(4)}.done`]: 'wrapper'
            });
            const summary = task.summarizeReviewActivity(stateDir);
            assert.equal(summary.completedPosts, 3);
            assert.equal(summary.unfinishedAttempts, 0);
        });

        it('does not count a done marker without a matching started marker as a post', () => {
            const stateDir = newStateDir();
            writeActivity(stateDir, { [`${id(7)}.done`]: 'thread' });
            assert.equal(task.summarizeReviewActivity(stateDir).completedPosts, 0);
        });

        it('counts attempts that never finished and agent error markers', () => {
            const stateDir = newStateDir();
            writeActivity(stateDir, {
                [`${id(5)}.started`]: 'thread',
                [`${id(6)}.started`]: 'wrapper',
                'claude-result.agent-error': 'error_max_turns',
                'unrelated.txt': 'ignored'
            });
            const summary = task.summarizeReviewActivity(stateDir);
            assert.equal(summary.unfinishedAttempts, 2);
            assert.equal(summary.agentErrors, 1);
            assert.equal(summary.completedPosts, 0);
        });
    });

    describe('prepareReviewStateFolders', () => {
        it('creates the activity and context folders and confirms they accept writes', () => {
            const stateDir = newStateDir();
            assert.equal(task.prepareReviewStateFolders(stateDir), true);
            assert.ok(fs.existsSync(path.join(stateDir, 'activity')));
            assert.ok(fs.existsSync(path.join(stateDir, 'context')));
            // The probe leaves nothing behind for the activity summary to misread
            assert.deepEqual(fs.readdirSync(path.join(stateDir, 'activity')), []);
        });

        it('reports a state folder that cannot hold the subfolders', () => {
            const notADirectory = path.join(newStateDir(), 'file');
            fs.writeFileSync(notADirectory, 'x');
            assert.equal(task.prepareReviewStateFolders(notADirectory), false);
        });
    });

    describe('collectContextDegradations', () => {
        it('returns the reported details, or nothing when all context loaded', () => {
            const stateDir = newStateDir();
            assert.deepEqual(task.collectContextDegradations(stateDir), []);

            fs.mkdirSync(path.join(stateDir, 'context'));
            fs.writeFileSync(path.join(stateDir, 'context', 'pr-threads.degraded'), 'Existing comment threads could not be retrieved.');
            fs.writeFileSync(path.join(stateDir, 'context', 'work-item-details.degraded'), '');
            assert.deepEqual(task.collectContextDegradations(stateDir),
                ['Existing comment threads could not be retrieved.', 'work-item-details']);
        });
    });

    describe('isSnapshotBound', () => {
        const merged = '1'.repeat(40);

        it('accepts a diff-only review of that commit', () => {
            assert.equal(task.isSnapshotBound({ currentSourceCommit: SHA_C, diffOnlyActive: true }), true);
        });

        it('requires the PR validation build commit to match when it is known', () => {
            assert.equal(task.isSnapshotBound({ currentSourceCommit: SHA_C, diffOnlyActive: false, pullRequestSourceCommitId: SHA_C.toUpperCase() }), true);
            // A push after checkout: the build reviewed an older commit
            assert.equal(task.isSnapshotBound({ currentSourceCommit: SHA_C, diffOnlyActive: false, pullRequestSourceCommitId: SHA_A, headCommit: SHA_C }), false);
        });

        it('otherwise requires the commit to be checked out or merged into the checkout', () => {
            assert.equal(task.isSnapshotBound({ currentSourceCommit: SHA_C, diffOnlyActive: false, headCommit: SHA_C }), true);
            assert.equal(task.isSnapshotBound({ currentSourceCommit: SHA_C, diffOnlyActive: false, headCommit: merged, headSecondParent: SHA_C }), true);
            // Manual run with another branch checked out
            assert.equal(task.isSnapshotBound({ currentSourceCommit: SHA_C, diffOnlyActive: false, headCommit: SHA_A }), false);
        });

        it('rejects a missing or malformed commit', () => {
            assert.equal(task.isSnapshotBound({ currentSourceCommit: '', diffOnlyActive: true }), false);
            assert.equal(task.isSnapshotBound({ currentSourceCommit: 'abc', diffOnlyActive: true }), false);
        });
    });

    describe('shouldRecordReviewState', () => {
        function recordableInput() {
            return {
                state: {
                    currentIteration: 3, currentSourceCommit: SHA_C, fullChangesStatus: 'Complete',
                    stateReadStatus: 'NotFound', lastReviewedIteration: null, lastReviewedSourceCommit: null,
                    delta: { status: 'NotApplicable', reason: null, files: [] }
                },
                activity: { readable: true, completedPosts: 2, unfinishedAttempts: 0, agentErrors: 0 },
                reviewContextWarnings: false,
                contextDegradations: [],
                snapshotBound: true,
                suppressPositiveFeedback: false
            };
        }

        it('records a complete review', () => {
            assert.equal(task.shouldRecordReviewState(recordableInput()).record, true);
        });

        const blockers = {
            'no review state': input => { input.state = null; },
            'a malformed current commit': input => { input.state.currentSourceCommit = 'abc'; },
            'unreadable activity': input => { input.activity.readable = false; },
            'an agent error result': input => { input.activity.agentErrors = 1; },
            'an unfinished posting attempt': input => { input.activity.unfinishedAttempts = 1; },
            'context fetch warnings': input => { input.reviewContextWarnings = true; },
            'a context degradation marker': input => { input.contextDegradations = ['Existing comment threads could not be retrieved.']; },
            'an incomplete changed-file list': input => { input.state.fullChangesStatus = 'Partial'; },
            'an unbound snapshot': input => { input.snapshotBound = false; },
            'zero posts with positive feedback allowed': input => { input.activity.completedPosts = 0; }
        };
        for (const [name, apply] of Object.entries(blockers)) {
            it(`does not record with ${name}`, () => {
                const input = recordableInput();
                apply(input);
                const decision = task.shouldRecordReviewState(input);
                assert.equal(decision.record, false);
                assert.ok(decision.reason.length > 0);
            });
        }

        it('records a clean review that posted nothing when positive feedback is suppressed', () => {
            const input = recordableInput();
            input.activity.completedPosts = 0;
            input.suppressPositiveFeedback = true;
            assert.equal(task.shouldRecordReviewState(input).record, true);
        });

        it('never records a run with a failed attempt, even with positive feedback suppressed', () => {
            const input = recordableInput();
            input.activity.completedPosts = 0;
            input.activity.unfinishedAttempts = 1;
            input.suppressPositiveFeedback = true;
            assert.equal(task.shouldRecordReviewState(input).record, false);
        });
    });

    describe('buildDeltaPathspecs', () => {
        it('builds literal, repository-relative pathspecs including rename origins', () => {
            const specs = task.buildDeltaPathspecs([
                { path: '/src/a.cs', changeType: 'edit' },
                { path: '/src/new [1].cs', changeType: 'rename', originalPath: '/src/old.cs' },
                { path: '/src/a.cs', changeType: 'edit' }
            ]);
            assert.deepEqual(specs, [':(literal)src/a.cs', ':(literal)src/new [1].cs', ':(literal)src/old.cs']);
        });

        it('returns null for no files or a list too long for a command line', () => {
            assert.equal(task.buildDeltaPathspecs([]), null);
            const many = Array.from({ length: 6 }, (_, i) => ({ path: `/f${i}.cs`, changeType: 'edit' }));
            assert.equal(task.buildDeltaPathspecs(many, { maxPaths: 5, maxChars: 20000 }), null);
            assert.equal(task.buildDeltaPathspecs(many, { maxPaths: 50, maxChars: 30 }), null);
        });
    });
});
