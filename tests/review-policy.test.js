/**
 * Tests for the review policy rules in index.ts (PR #65 rework).
 *
 * The task writes review-policy.json before the review agent runs; the comment
 * scripts enforce it (tests/powershell/ReviewPolicy.Tests.ps1). The rules for
 * which threads the agent may act on live here, so they are tested as a truth
 * table over thread authorship and status, resolveThreads, and template vs. raw
 * prompts.
 *
 * Requires the compiled task (`npm run build`); importing index.js is side-effect free.
 */

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const compiledTask = path.join(__dirname, '..', 'CopilotCodeReviewV1', 'index.js');

describe('review policy (PR #65 rework)', () => {
    let task;
    let tmpDir;

    before(() => {
        assert.ok(fs.existsSync(compiledTask),
            'CopilotCodeReviewV1/index.js not found — run `npm run build` before `npm test`.');
        task = require(compiledTask);
        tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'review-policy-'));
    });

    after(() => {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    const thread = (threadId, agentAuthored, status, comments = []) => ({ threadId, agentAuthored, status, comments });

    function policyInput(overrides = {}) {
        return {
            collectionUri: 'https://dev.azure.com/org',
            project: 'Proj',
            repository: 'Repo',
            pullRequestId: 5,
            resolveThreads: 'all',
            rawPrompt: false,
            workspaceRoot: path.resolve('/work/s'),
            threads: [],
            reviewState: { fullChangesStatus: 'Complete', changedFiles: ['/src/a.cs'] },
            ...overrides
        };
    }

    describe('buildReviewPolicy', () => {
        // Every kind of thread, keyed by a label used in the table below
        const threads = [
            thread(1, true, 'active'),
            thread(2, true, 'fixed'),
            thread(3, false, 'active'),
            thread(4, false, 'pending'),
            thread(5, false, 'fixed'),
            thread(6, false, 'closed'),
            thread(7, false, 'Active')
        ];
        // Thread ids the agent may reply to and resolve
        const table = [
            { resolveThreads: 'agentOnly', rawPrompt: false, allowed: [1, 2] },
            { resolveThreads: 'all', rawPrompt: false, allowed: [1, 2, 3, 4, 7] },
            // Raw prompts: resolveThreads does not apply (documented contract)
            { resolveThreads: 'agentOnly', rawPrompt: true, allowed: [1, 2, 3, 4, 5, 6, 7] },
            { resolveThreads: 'all', rawPrompt: true, allowed: [1, 2, 3, 4, 5, 6, 7] }
        ];

        for (const row of table) {
            it(`allows threads ${row.allowed.join(', ')} with resolveThreads '${row.resolveThreads}' and a ${row.rawPrompt ? 'raw' : 'template'} prompt`, () => {
                const policy = task.buildReviewPolicy(policyInput({ threads, resolveThreads: row.resolveThreads, rawPrompt: row.rawPrompt }));
                assert.deepEqual(policy.threads.filter(t => t.mayReply).map(t => t.threadId), row.allowed);
                // Replying and resolving go together
                for (const entry of policy.threads) {
                    assert.equal(entry.mayUpdateStatus, entry.mayReply, `thread ${entry.threadId}`);
                }
            });
        }

        it('limits edits and deletions to comments an automated review wrote, in every mode', () => {
            const withComments = [
                thread(1, true, 'active', [{ commentId: 1, agentAuthored: true }, { commentId: 2, agentAuthored: false }]),
                thread(3, false, 'active', [{ commentId: 1, agentAuthored: false }, { commentId: 4, agentAuthored: true }])
            ];
            for (const rawPrompt of [false, true]) {
                const policy = task.buildReviewPolicy(policyInput({ threads: withComments, rawPrompt }));
                assert.deepEqual(policy.threads.map(t => t.agentCommentIds), [[1], [4]]);
            }
        });

        it('allows no thread actions when the threads could not be retrieved', () => {
            const policy = task.buildReviewPolicy(policyInput({ threads: null, rawPrompt: true }));
            assert.equal(policy.threadsComplete, false);
            assert.deepEqual(policy.threads, []);
        });

        it('carries the destination, the workspace, and the prompt settings', () => {
            const policy = task.buildReviewPolicy(policyInput());
            assert.equal(policy.schemaVersion, 1);
            assert.deepEqual(policy.destination, {
                collectionUri: 'https://dev.azure.com/org', project: 'Proj', repository: 'Repo', pullRequestId: 5
            });
            assert.equal(policy.workspaceRoot, path.resolve('/work/s'));
            assert.equal(policy.resolveThreads, 'all');
            assert.equal(policy.rawPrompt, false);
            assert.equal(policy.threadsComplete, true);
        });

        it('marks the changed-file list complete only when the full list was retrieved', () => {
            const complete = task.buildReviewPolicy(policyInput());
            assert.equal(complete.changedFilesComplete, true);
            assert.deepEqual(complete.changedFiles, ['/src/a.cs']);

            const partial = task.buildReviewPolicy(policyInput({ reviewState: { fullChangesStatus: 'Partial', changedFiles: ['/src/a.cs'] } }));
            assert.equal(partial.changedFilesComplete, false);
            assert.deepEqual(partial.changedFiles, ['/src/a.cs']);

            const unreported = task.buildReviewPolicy(policyInput({ reviewState: { fullChangesStatus: 'Complete', changedFiles: null } }));
            assert.equal(unreported.changedFilesComplete, false);
            assert.deepEqual(unreported.changedFiles, []);

            const missing = task.buildReviewPolicy(policyInput({ reviewState: null }));
            assert.equal(missing.changedFilesComplete, false);
            assert.deepEqual(missing.changedFiles, []);
        });
    });

    describe('readPolicyThreadData', () => {
        function write(contents) {
            const file = path.join(fs.mkdtempSync(path.join(tmpDir, 'threads-')), 'PR_Threads.json');
            fs.writeFileSync(file, contents, 'utf8');
            return file;
        }

        it('reads the threads and drops unusable entries', () => {
            const data = task.readPolicyThreadData(write('﻿' + JSON.stringify({
                schemaVersion: 1,
                threadsStatus: 'Complete',
                threads: [
                    { threadId: 31, status: 'active', agentAuthored: true, comments: [{ commentId: 1, agentAuthored: true }, { commentId: 0 }, null] },
                    { threadId: 'x', status: 'active' },
                    { threadId: 32, status: 'fixed', agentAuthored: 'yes' }
                ]
            })));
            assert.deepEqual(data, [
                { threadId: 31, status: 'active', agentAuthored: true, comments: [{ commentId: 1, agentAuthored: true }] },
                // Anything but a literal true is not agent-authored
                { threadId: 32, status: 'fixed', agentAuthored: false, comments: [] }
            ]);
        });

        it('returns null for failed retrieval or a missing or malformed file', () => {
            assert.equal(task.readPolicyThreadData(write(JSON.stringify({ threadsStatus: 'Failed', threads: [] }))), null);
            assert.equal(task.readPolicyThreadData(write('not json')), null);
            assert.equal(task.readPolicyThreadData(path.join(tmpDir, 'missing.json')), null);
        });
    });

    describe('describePolicyRefusals', () => {
        it('lists each refusal, then caps the list', () => {
            assert.deepEqual(task.describePolicyRefusals([]), []);
            const reasons = Array.from({ length: 12 }, (_, i) => `reason ${i + 1}`);
            const messages = task.describePolicyRefusals(reasons);
            assert.equal(messages.length, 11);
            assert.equal(messages[0], 'The review policy refused an action by the review agent: reason 1');
            assert.match(messages[10], /refused 2 more action\(s\)/);
        });
    });
});
