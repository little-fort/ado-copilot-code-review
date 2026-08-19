import * as tl from 'azure-pipelines-task-lib/task';
import * as path from 'path';
import * as fs from 'fs';
import * as child_process from 'child_process';
import * as os from 'os';

/**
 * Check if PowerShell 7 (pwsh) is available on the system
 */
function checkPwshAvailable(): boolean {
    try {
        const result = child_process.spawnSync('pwsh', ['--version'], {
            encoding: 'utf8',
            shell: true
        });
        return result.status === 0;
    } catch {
        return false;
    }
}

/**
 * Check if we're running on Windows
 */
function isWindows(): boolean {
    return process.platform === 'win32';
}

/**
 * Neutralize Azure Pipelines logging-command markers in untrusted text before logging.
 * Prepends a space to any line starting with `##` so Azure Pipelines won't interpret
 * `##vso[...]` or `##[...]` patterns as logging commands.
 */
function sanitizePipelineLog(text: string): string {
    return text.replace(/^##/gm, ' ##');
}

/**
 * Normalizes a GitHub Enterprise host input to a bare hostname.
 * GH_HOST expects a hostname (e.g. 'subdomain.ghe.com'), but users commonly
 * paste a full URL — strip any scheme and trailing slashes.
 */
function normalizeGitHubHost(host: string): string {
    return host.trim().replace(/^https?:\/\//i, '').replace(/\/+$/, '');
}

/**
 * Builds the "Review Behavior Directives" prompt section from task inputs
 * (issues #55 and #28). The section is appended to template-based prompts for
 * both Copilot and Claude Code; raw prompts are contractually passed as-is
 * and never receive it.
 */
function buildReviewDirectives(resolveThreads: string, suppressPositiveFeedback: boolean): string {
    const directives: string[] = [];

    if (resolveThreads === 'all') {
        directives.push(
            'Thread resolution scope: In addition to the automated-review threads listed in the ' +
            'COPILOT COMMENT THREADS (JSON) section of PR_Details.txt, you may resolve threads created by ' +
            'human reviewers when the changes in the current iteration clearly and completely address their ' +
            'concern. Follow the same two-step process: first reply to the thread explaining what change ' +
            "addressed it, then mark the thread as fixed. If in doubt, leave human reviewers' threads untouched.");
    } else {
        directives.push(
            'Thread resolution scope: Only resolve or change the status of comment threads that appear in ' +
            'the COPILOT COMMENT THREADS (JSON) section of PR_Details.txt — those threads were created by ' +
            'previous automated reviews. NEVER resolve, close, or otherwise change the status of threads ' +
            'created by human reviewers, even if the underlying issue appears to have been addressed; leave ' +
            'those threads to the people involved in them.');
    }

    if (suppressPositiveFeedback) {
        directives.push(
            "Positive feedback suppression: Do NOT post positive, congratulatory, or 'looks good' comments. " +
            'If the review finds no impactful issues, do not post any comment at all — simply end the review. ' +
            'This overrides any earlier instruction to leave a comment indicating the code looks good to ' +
            'merge. Only post comments that request changes or flag genuine concerns.');
    }

    return '\n\n# Review Behavior Directives\n\n' +
        'The following directives are configured by the pipeline administrator and take precedence over any ' +
        'conflicting instructions elsewhere in this prompt:\n\n' +
        directives.map(d => '- ' + d).join('\n\n') + '\n';
}

async function run(): Promise<void> {
    try {
        // Check prerequisites first
        console.log('Checking prerequisites...');
        
        // Check if PowerShell 7 (pwsh) is available
        if (!checkPwshAvailable()) {
            tl.setResult(tl.TaskResult.Failed, 
                'PowerShell 7 (pwsh) is required but not found. ' +
                'Please install PowerShell 7 or later. ' +
                'Visit https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell for installation instructions.');
            return;
        }
        console.log('PowerShell 7 (pwsh) is available.');

        // Check author filter first (before any other processing)
        const authors = tl.getInput('authors');
        if (authors) {
            const requestedForEmail = tl.getVariable('Build.RequestedForEmail') || '';
            const authorList = authors.split(',').map(email => email.trim().toLowerCase());
            const currentAuthor = requestedForEmail.toLowerCase();
            
            console.log('='.repeat(60));
            console.log('Author Filter Check');
            console.log('='.repeat(60));
            console.log(`Configured authors: ${authorList.join(', ')}`);
            console.log(`PR author email: ${requestedForEmail || '(not available)'}`);
            
            if (!authorList.includes(currentAuthor)) {
                console.log(`Result: PR author is NOT in the configured authors list.`);
                console.log('Skipping code review for this PR.');
                console.log('='.repeat(60));
                tl.setResult(tl.TaskResult.Succeeded, 'Skipped: PR author not in configured authors list.');
                return;
            }
            
            console.log(`Result: PR author IS in the configured authors list.`);
            console.log('Proceeding with code review.');
            console.log('='.repeat(60));
        }

        // Get agent selection and auth inputs
        const useClaudeCode = tl.getBoolInput('useClaudeCode', false);
        const useCustomModelProvider = tl.getBoolInput('useCustomModelProvider', false);
        const githubPat = tl.getInput('githubPat');
        const githubHost = tl.getInput('githubHost');
        const anthropicApiKey = tl.getInput('anthropicApiKey');
        const claudeCodeOAuthToken = tl.getInput('claudeCodeOAuthToken');
        const maxTurns = tl.getInput('maxTurns');
        const maxBudget = tl.getInput('maxBudget');
        const copilotProviderBaseUrl = tl.getInput('copilotProviderBaseUrl');
        const copilotProviderType = tl.getInput('copilotProviderType');
        const copilotProviderApiKey = tl.getInput('copilotProviderApiKey');
        const model = tl.getInput('model');

        // Validate agent-specific auth
        if (useClaudeCode && useCustomModelProvider) {
            tl.setResult(tl.TaskResult.Failed,
                "The 'useClaudeCode' and 'useCustomModelProvider' inputs are mutually exclusive—custom model providers " +
                'apply only to the GitHub Copilot CLI. Please enable at most one of them.');
            return;
        }

        if (useClaudeCode) {
            if (anthropicApiKey && claudeCodeOAuthToken) {
                tl.setResult(tl.TaskResult.Failed,
                    "Both 'anthropicApiKey' and 'claudeCodeOAuthToken' are set. Please provide exactly one — " +
                    'the API key bills against Anthropic API usage, while the OAuth token bills against a Claude subscription.');
                return;
            }
            if (!anthropicApiKey && !claudeCodeOAuthToken) {
                tl.setResult(tl.TaskResult.Failed,
                    'Claude Code CLI requires authentication. Provide either the anthropicApiKey input (Anthropic API billing) ' +
                    "or the claudeCodeOAuthToken input (Claude subscription billing — generate one by running 'claude setup-token' locally).");
                return;
            }
        } else if (useCustomModelProvider) {
            // BYOK mode: the Copilot CLI talks directly to the configured provider,
            // so a GitHub PAT is optional (it only adds GitHub-backed features like code search)
            if (!copilotProviderBaseUrl) {
                tl.setResult(tl.TaskResult.Failed,
                    'A provider base URL is required when using a custom model provider. Please provide the copilotProviderBaseUrl input.');
                return;
            }
            if (!copilotProviderType) {
                tl.setResult(tl.TaskResult.Failed,
                    'A provider type is required when using a custom model provider. Please provide the copilotProviderType input.');
                return;
            }
            if (!model) {
                tl.setResult(tl.TaskResult.Failed,
                    'A model is required when using a custom model provider. Please set the model input to a model identifier your provider serves.');
                return;
            }
            const knownProviderTypes = ['openai', 'azure', 'anthropic'];
            if (!knownProviderTypes.includes(copilotProviderType.toLowerCase())) {
                tl.warning(`Provider type '${copilotProviderType}' is not one of the documented values (${knownProviderTypes.join(', ')}). ` +
                    'The Copilot CLI may reject it.');
            }
            if (!copilotProviderApiKey) {
                console.log('No custom provider API key supplied—assuming the provider endpoint does not require authentication.');
            }
        } else {
            if (!githubPat) {
                tl.setResult(tl.TaskResult.Failed,
                    'GitHub PAT is required when using GitHub Copilot CLI. Please provide the githubPat input.');
                return;
            }
        }

        // Get Azure DevOps authentication settings
        const useSystemAccessToken = tl.getBoolInput('useSystemAccessToken', false);
        const azureDevOpsPat = tl.getInput('azureDevOpsPat');
        
        // Determine which token and auth type to use
        let azureDevOpsToken: string;
        let azureDevOpsAuthType: string;
        
        if (useSystemAccessToken) {
            // Use System.AccessToken (OAuth Bearer token)
            const systemToken = tl.getVariable('System.AccessToken');
            if (!systemToken) {
                tl.setResult(tl.TaskResult.Failed, 
                    'System.AccessToken is not available. Ensure the pipeline has access to the OAuth token. ' +
                    'In YAML pipelines, you may need to explicitly map it using env: SYSTEM_ACCESSTOKEN: $(System.AccessToken)');
                return;
            }
            azureDevOpsToken = systemToken;
            azureDevOpsAuthType = 'Bearer';
            console.log('Using System.AccessToken (OAuth) for Azure DevOps authentication.');
        } else if (azureDevOpsPat) {
            // Use provided PAT (Basic auth)
            azureDevOpsToken = azureDevOpsPat;
            azureDevOpsAuthType = 'Basic';
            console.log('Using Personal Access Token for Azure DevOps authentication.');
        } else {
            tl.setResult(tl.TaskResult.Failed, 
                'Azure DevOps authentication is required. Either provide an Azure DevOps PAT or enable "Use System Access Token".');
            return;
        }
        
        // Get inputs with defaults from pipeline variables
        let organization = tl.getInput('organization');
        let collectionUri = tl.getInput('collectionUri');
        let project = tl.getInput('project');
        let repository = tl.getInput('repository');

        // Resolve the collection URI using priority chain:
        // 1. Explicit collectionUri input
        // 2. organization input -> construct https://dev.azure.com/{org}
        // 3. System.CollectionUri pipeline variable (used directly)
        let resolvedCollectionUri: string | undefined;

        if (collectionUri) {
            resolvedCollectionUri = collectionUri.replace(/\/+$/, '');
            console.log(`Using explicit collection URI: ${resolvedCollectionUri}`);
        } else if (organization) {
            resolvedCollectionUri = `https://dev.azure.com/${organization}`;
            console.log(`Constructed collection URI from organization: ${resolvedCollectionUri}`);
        } else {
            const systemCollectionUri = tl.getVariable('System.CollectionUri');
            if (systemCollectionUri) {
                resolvedCollectionUri = systemCollectionUri.replace(/\/+$/, '');
                console.log(`Auto-detected collection URI from System.CollectionUri: ${resolvedCollectionUri}`);
            }
        }

        if (!resolvedCollectionUri) {
            tl.setResult(tl.TaskResult.Failed,
                'Collection URI could not be determined. Provide collectionUri, organization, or ensure System.CollectionUri is available.');
            return;
        }

        if (!project) {
            tl.setResult(tl.TaskResult.Failed, 'Project is required. Either provide it as an input or ensure System.TeamProject is available.');
            return;
        }

        if (!repository) {
            tl.setResult(tl.TaskResult.Failed, 'Repository is required. Either provide it as an input or ensure Build.Repository.Name is available.');
            return;
        }

        // Get optional inputs
        let pullRequestId = tl.getInput('pullRequestId');
        const timeoutMinutes = parseInt(tl.getInput('timeout') || '15', 10);
        const promptFile = tl.getInput('promptFile');
        const prompt = tl.getInput('prompt');
        const promptRaw = tl.getInput('promptRaw');
        const promptFileRaw = tl.getInput('promptFileRaw');
        const includeWorkItems = tl.getBoolInput('includeWorkItems', false);
        const diffOnlyReview = tl.getBoolInput('diffOnlyReview', false);
        const publishPromptArtifacts = tl.getBoolInput('publishPromptArtifacts', false);
        const resolveThreads = tl.getInput('resolveThreads') || 'agentOnly';
        const suppressPositiveFeedback = tl.getBoolInput('suppressPositiveFeedback', false);

        // Fail fast on a typo'd value — falling back silently would change
        // which PR threads the agent is allowed to resolve
        if (resolveThreads !== 'agentOnly' && resolveThreads !== 'all') {
            tl.setResult(tl.TaskResult.Failed,
                `Invalid resolveThreads value '${resolveThreads}'. Valid values: 'agentOnly', 'all'.`);
            return;
        }

        // filePath inputs return the working directory path when not set, so check both
        // the input value and that the path actually points to a real file.
        const isPromptFileSet = !!(promptFile && fs.existsSync(promptFile) && fs.statSync(promptFile).isFile());
        const isPromptFileRawSet = !!(promptFileRaw && fs.existsSync(promptFileRaw) && fs.statSync(promptFileRaw).isFile());

        // If PR ID not provided, try to get from pipeline variable
        if (!pullRequestId) {
            pullRequestId = tl.getVariable('System.PullRequest.PullRequestId');
        }

        if (!pullRequestId) {
            tl.setResult(tl.TaskResult.Failed, 'Pull Request ID is required. Either provide it as an input or run this task as part of a PR validation build.');
            return;
        }

        const agentName = useClaudeCode ? 'Claude Code' : 'Copilot';
        console.log('='.repeat(60));
        console.log(`${agentName} Code Review Task`);
        console.log('='.repeat(60));
        console.log(`Agent: ${agentName}`);
        if (useCustomModelProvider) {
            console.log(`Model Provider: custom (${copilotProviderType})`);
        }
        console.log(`Collection URI: ${resolvedCollectionUri}`);
        console.log(`Project: ${project}`);
        console.log(`Repository: ${repository}`);
        console.log(`Pull Request ID: ${pullRequestId}`);
        console.log(`Timeout: ${timeoutMinutes} minutes`);
        if (model) {
            console.log(`Model: ${model}`);
        }
        console.log('='.repeat(60));

        // Set environment variables for the CLI agent
        if (useClaudeCode) {
            if (claudeCodeOAuthToken) {
                // Long-lived subscription token generated via `claude setup-token`
                process.env['CLAUDE_CODE_OAUTH_TOKEN'] = claudeCodeOAuthToken;
                console.log('Using Claude Code OAuth token (subscription billing) for authentication.');
            } else {
                process.env['ANTHROPIC_API_KEY'] = anthropicApiKey!;
                console.log('Using Anthropic API key (API usage billing) for authentication.');
            }
        } else {
            if (useCustomModelProvider) {
                // The Copilot CLI reads these to route model requests to a BYOK provider
                // instead of GitHub-hosted models
                process.env['COPILOT_PROVIDER_BASE_URL'] = copilotProviderBaseUrl!;
                process.env['COPILOT_PROVIDER_TYPE'] = copilotProviderType!;
                if (copilotProviderApiKey) {
                    process.env['COPILOT_PROVIDER_API_KEY'] = copilotProviderApiKey;
                }
            }
            // Optional in BYOK mode: a PAT additionally enables GitHub-backed features (e.g. code search)
            if (githubPat) {
                process.env['GH_TOKEN'] = githubPat;
            }
            // GitHub Enterprise (e.g. subdomain.ghe.com): the Copilot CLI honors
            // GH_HOST when validating the PAT, so enterprise-scoped tokens
            // authenticate against the right instance (issue #59)
            if (githubHost) {
                const normalizedHost = normalizeGitHubHost(githubHost);
                process.env['GH_HOST'] = normalizedHost;
                console.log(`Using GitHub Enterprise host: ${normalizedHost}`);
            }
        }
        process.env['AZUREDEVOPS_TOKEN'] = azureDevOpsToken;
        process.env['AZUREDEVOPS_AUTH_TYPE'] = azureDevOpsAuthType;
        process.env['AZUREDEVOPS_COLLECTION_URI'] = resolvedCollectionUri;
        process.env['PROJECT'] = project;
        process.env['REPOSITORY'] = repository;
        process.env['PRID'] = pullRequestId;

        const scriptsDir = path.join(__dirname, 'scripts');
        const workingDirectory = tl.getVariable('System.DefaultWorkingDirectory') || process.cwd();

        // On Windows, refresh PATH from the registry so we can detect CLIs installed
        // after the agent service started (the agent inherits its PATH at service startup)
        if (isWindows()) {
            try {
                const refreshResult = child_process.spawnSync('pwsh', [
                    '-NoProfile', '-Command',
                    `[System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")`
                ], { encoding: 'utf8', shell: false });
                if (refreshResult.status === 0 && refreshResult.stdout.trim()) {
                    // Merge registry PATH with existing process PATH to preserve
                    // Azure Pipelines job-scoped entries (tool cache, task-prepended paths)
                    const registryPath = refreshResult.stdout.trim();
                    const currentPath = process.env['PATH'] || '';
                    const seen = new Set<string>();
                    const merged: string[] = [];
                    for (const entry of registryPath.split(';').concat(currentPath.split(';'))) {
                        const trimmed = entry.trim();
                        if (trimmed && !seen.has(trimmed.toLowerCase())) {
                            seen.add(trimmed.toLowerCase());
                            merged.push(trimmed);
                        }
                    }
                    process.env['PATH'] = merged.join(';');
                }
            } catch {
                // Non-fatal — continue with existing PATH
            }
        }

        // Step 1: Install CLI agent if not present
        if (useClaudeCode) {
            console.log('\n[Step 1/5] Checking Claude Code CLI installation...');
            const claudeInstalled = await checkClaudeCodeCli();
            if (!claudeInstalled) {
                console.log('Claude Code CLI not found. Installing...');
                await installClaudeCodeCli();
            } else {
                console.log('Claude Code CLI is already installed.');
            }
        } else {
            console.log('\n[Step 1/5] Checking GitHub Copilot CLI installation...');
            const copilotInstalled = await checkCopilotCli();
            if (!copilotInstalled) {
                console.log('GitHub Copilot CLI not found. Installing...');
                await installCopilotCli();
            } else {
                console.log('GitHub Copilot CLI is already installed.');
            }
        }

        // Step 2: Fetch PR details
        console.log('\n[Step 2/5] Fetching pull request details...');
        const prDetailsScript = path.join(scriptsDir, 'Get-AzureDevOpsPR.ps1');
        const prDetailsOutput = path.join(workingDirectory, 'PR_Details.txt');
        
        await runPowerShellScript(prDetailsScript, [
            `-Token "${azureDevOpsToken}"`,
            `-AuthType "${azureDevOpsAuthType}"`,
            `-CollectionUri "${resolvedCollectionUri}"`,
            `-Project "${project}"`,
            `-Repository "${repository}"`,
            `-Id ${pullRequestId}`,
            `-OutputFile "${prDetailsOutput}"`
        ]);
        console.log(`PR details saved to: ${prDetailsOutput}`);

        // Step 3: Fetch PR changes (iteration details)
        console.log('\n[Step 3/5] Fetching pull request changes...');
        const prChangesScript = path.join(scriptsDir, 'Get-AzureDevOpsPRChanges.ps1');
        const iterationDetailsOutput = path.join(workingDirectory, 'Iteration_Details.txt');
        
        await runPowerShellScript(prChangesScript, [
            `-Token "${azureDevOpsToken}"`,
            `-AuthType "${azureDevOpsAuthType}"`,
            `-CollectionUri "${resolvedCollectionUri}"`,
            `-Project "${project}"`,
            `-Repository "${repository}"`,
            `-Id ${pullRequestId}`,
            `-OutputFile "${iterationDetailsOutput}"`
        ]);
        console.log(`Iteration details saved to: ${iterationDetailsOutput}`);

        // Read the iteration ID from the file written by Get-AzureDevOpsPRChanges.ps1
        const iterationIdFile = path.join(workingDirectory, 'Iteration_Id.txt');
        if (fs.existsSync(iterationIdFile)) {
            const iterationId = fs.readFileSync(iterationIdFile, 'utf8').trim();
            if (iterationId) {
                process.env['ITERATION_ID'] = iterationId;
                console.log(`Iteration ID set to: ${iterationId}`);
            }
        }

        // Diff-only mode: pre-compute the git diff from commit SHAs
        let diffContent: string | null = null;

        if (diffOnlyReview) {
            console.log('\n[Diff-Only Mode] Computing PR diff from commit SHAs...');
            const sourceCommitFile = path.join(workingDirectory, 'Source_Commit.txt');
            const targetCommitFile = path.join(workingDirectory, 'Target_Commit.txt');

            if (!fs.existsSync(sourceCommitFile) || !fs.existsSync(targetCommitFile)) {
                tl.setResult(tl.TaskResult.Failed,
                    'Diff-only mode failed: Commit SHA files (Source_Commit.txt / Target_Commit.txt) not found. ' +
                    'This may indicate the Azure DevOps API did not return commit references for the PR iteration.');
                return;
            }

            const sourceCommit = fs.readFileSync(sourceCommitFile, 'utf8').trim();
            const targetCommit = fs.readFileSync(targetCommitFile, 'utf8').trim();

            if (!sourceCommit || !targetCommit) {
                tl.setResult(tl.TaskResult.Failed,
                    'Diff-only mode failed: Commit SHA files are empty. ' +
                    'Source_Commit.txt and Target_Commit.txt must contain valid commit hashes.');
                return;
            }

            console.log(`  Source commit: ${sourceCommit}`);
            console.log(`  Target commit: ${targetCommit}`);

            try {
                // Three-dot syntax (target...source) diffs from the merge-base of the two
                // commits to source. This matches the PR-style diff shown in the ADO UI
                // and excludes unrelated target-branch changes when source is behind target.
                const diffResult = child_process.spawnSync(
                    'git',
                    ['diff', `${targetCommit}...${sourceCommit}`],
                    {
                        encoding: 'utf8',
                        cwd: workingDirectory,
                        maxBuffer: 50 * 1024 * 1024
                    }
                );

                if (diffResult.status === 0 && diffResult.stdout.trim()) {
                    diffContent = diffResult.stdout;
                    console.log(`  Diff computed successfully (${diffContent.length} characters).`);
                } else if (diffResult.status === 0 && !diffResult.stdout.trim()) {
                    tl.setResult(tl.TaskResult.Failed,
                        'Diff-only mode failed: git diff returned empty output. ' +
                        'The PR may have no code changes between the source and target commits.');
                    return;
                } else {
                    const stderr = diffResult.stderr ? diffResult.stderr.trim() : '';
                    tl.setResult(tl.TaskResult.Failed,
                        `Diff-only mode failed: git diff exited with code ${diffResult.status}. ` +
                        'This typically means the commit SHAs are not available in the local checkout. ' +
                        'Ensure your pipeline uses "fetchDepth: 0" for a full clone. ' +
                        (stderr ? `git stderr: ${stderr}` : ''));
                    return;
                }
            } catch (err) {
                tl.setResult(tl.TaskResult.Failed,
                    `Diff-only mode failed: git diff threw an error: ${err instanceof Error ? err.message : String(err)}. ` +
                    'Ensure git is available and the working directory is a valid repository.');
                return;
            }
        }

        // Step 4: Fetch linked work item details (optional)
        if (includeWorkItems) {
            console.log('\n[Step 4/5] Fetching linked work item details...');
            const workItemIdsFile = path.join(workingDirectory, 'Work_Item_Ids.txt');

            if (fs.existsSync(workItemIdsFile)) {
                const workItemIds = fs.readFileSync(workItemIdsFile, 'utf8').trim();

                if (workItemIds) {
                    const workItemsScript = path.join(scriptsDir, 'Get-AzureDevOpsWorkItems.ps1');
                    const workItemDetailsOutput = path.join(workingDirectory, 'Work_Item_Details.txt');

                    try {
                        await runPowerShellScript(workItemsScript, [
                            `-Token "${azureDevOpsToken}"`,
                            `-AuthType "${azureDevOpsAuthType}"`,
                            `-CollectionUri "${resolvedCollectionUri}"`,
                            `-Project "${project}"`,
                            `-WorkItemIds "${workItemIds}"`,
                            `-OutputFile "${workItemDetailsOutput}"`
                        ]);
                        console.log(`Work item details saved to: ${workItemDetailsOutput}`);
                    } catch (err) {
                        console.log('Warning: Failed to fetch work item details. Continuing without work item context.');
                        console.log(`Error: ${err instanceof Error ? err.message : String(err)}`);
                    }
                } else {
                    console.log('No linked work item IDs found. Skipping work item detail fetch.');
                }
            } else {
                console.log('No linked work items for this PR. Skipping work item detail fetch.');
            }
        } else {
            console.log('\n[Step 4/5] Skipping work item details (disabled).');
        }

        // Step 5: Run CLI agent for code review
        console.log(`\n[Step 5/5] Running ${agentName} code review...`);
        
        // Determine the prompt file to use
        let promptFilePath: string = '';
        let customPromptText: string | null = null;

        // Validate that only one prompt input is provided
        const activePromptInputs: string[] = [];
        if (prompt) activePromptInputs.push('prompt');
        if (isPromptFileSet) activePromptInputs.push('promptFile');
        if (promptRaw) activePromptInputs.push('promptRaw');
        if (isPromptFileRawSet) activePromptInputs.push('promptFileRaw');

        if (activePromptInputs.length > 1) {
            tl.setResult(tl.TaskResult.Failed,
                `Multiple prompt inputs are set (${activePromptInputs.join(', ')}). Only one prompt input should be provided. ` +
                'Please use only one of: prompt, promptFile, promptRaw, or promptFileRaw.');
            return;
        }

        if (promptRaw) {
            // Raw prompt: pass directly to CLI with no modification
            console.log('Using raw prompt from input.');
            promptFilePath = path.join(workingDirectory, '_copilot_prompt.txt');
            fs.writeFileSync(promptFilePath, promptRaw, 'utf8');
            console.log('\nRAW PROMPT:\n' + sanitizePipelineLog(promptRaw) + '\n\n');
        } else if (isPromptFileRawSet) {
            // Raw prompt file: use file contents as-is with no modification
            console.log(`Using raw prompt from file: ${promptFileRaw}`);
            const fileContent = fs.readFileSync(promptFileRaw!, 'utf8');
            if (!fileContent.trim()) {
                tl.setResult(tl.TaskResult.Failed, `Raw prompt file is empty: ${promptFileRaw}`);
                return;
            }
            promptFilePath = path.join(workingDirectory, '_copilot_prompt.txt');
            fs.writeFileSync(promptFilePath, fileContent, 'utf8');
            console.log('\nRAW PROMPT:\n' + sanitizePipelineLog(fileContent) + '\n\n');
        } else if (prompt) {
            // Direct prompt input: merge with template
            console.log('Using custom prompt from input.');
            customPromptText = prompt;
        } else if (isPromptFileSet) {
            // Read from prompt file: merge with template
            console.log(`Using custom prompt from file: ${promptFile}`);
            const fileContent = fs.readFileSync(promptFile!, 'utf8').trim();
            if (!fileContent) {
                tl.setResult(tl.TaskResult.Failed, `Prompt file is empty: ${promptFile}`);
                return;
            }
            customPromptText = fileContent;
        }

        if (customPromptText) {
            // Use custom prompt template with placeholder replacement
            const customPromptTemplate = path.join(scriptsDir, 'prompt-custom.txt');
            const templateContent = fs.readFileSync(customPromptTemplate, 'utf8');
            const mergedPrompt = templateContent.replace('%CUSTOMPROMPT%', customPromptText);
            console.log('\nCUSTOM PROMPT:\n' + sanitizePipelineLog(mergedPrompt) + '\n\n');

            // Write merged prompt to a temp file in the working directory
            promptFilePath = path.join(workingDirectory, '_copilot_prompt.txt');
            fs.writeFileSync(promptFilePath, mergedPrompt, 'utf8');
            console.log('Custom prompt merged with instruction template.');
        } else if (!promptRaw && !isPromptFileRawSet) {
            // Use default prompt file bundled with the task
            promptFilePath = path.join(scriptsDir, 'prompt.txt');
            console.log('Using default prompt.');
        }

        // Append administrator-configured behavior directives (thread resolution
        // scope, positive feedback suppression) to template-based prompts.
        // Raw prompts are contractually passed as-is, so they are never modified.
        if (!promptRaw && !isPromptFileRawSet && promptFilePath) {
            const promptWithDirectives = fs.readFileSync(promptFilePath, 'utf8')
                + buildReviewDirectives(resolveThreads, suppressPositiveFeedback);
            const directivePromptPath = path.join(workingDirectory, '_directive_prompt.txt');
            fs.writeFileSync(directivePromptPath, promptWithDirectives, 'utf8');
            promptFilePath = directivePromptPath;
            console.log(`Applied review behavior directives (resolveThreads: ${resolveThreads}, suppressPositiveFeedback: ${suppressPositiveFeedback}).`);
        }

        // When using Claude Code, replace the Copilot attribution tag in the prompt
        if (useClaudeCode && promptFilePath) {
            let promptContent = fs.readFileSync(promptFilePath, 'utf8');
            if (promptContent.includes('Generated by GitHub Copilot')) {
                promptContent = promptContent.replace(/Generated by GitHub Copilot/g, 'Generated by Claude Code');
                // Always write to a separate temp file to avoid mutating the original prompt file
                const modifiedPromptPath = path.join(workingDirectory, '_claude_prompt.txt');
                fs.writeFileSync(modifiedPromptPath, promptContent, 'utf8');
                promptFilePath = modifiedPromptPath;
                console.log('Replaced attribution tag for Claude Code in prompt.');
            }
        }

        // Diff-only mode: inject context + diff into the prompt and flag for tool restriction
        let diffOnlyActive = false;

        if (diffOnlyReview && diffContent && promptFilePath) {
            let currentPrompt = fs.readFileSync(promptFilePath, 'utf8');
            const isRawPrompt = !!(promptRaw || isPromptFileRawSet);

            // For non-raw prompts (default/custom templates), attempt to rewrite instructions
            // that tell the agent to use git/file access. These are best-effort — the override
            // block appended at the end guarantees correct behavior regardless.
            if (!isRawPrompt) {
                const originalGuidelines = 'Using this information along with git commands and the local copy of the repository, please conduct a code review of this pull request and identify any suggested modifications that should be made.';
                const replacedGuidelines = 'Using the PR details and code diff provided at the end of this prompt, please conduct a code review of this pull request and identify any suggested modifications that should be made.';
                if (currentPrompt.includes(originalGuidelines)) {
                    currentPrompt = currentPrompt.replace(originalGuidelines, replacedGuidelines);
                }

                const originalOverview = 'The details of a pull request for the repo in the working directory have been saved to the PR_Details.txt file—please review this file for broad context on the pull request. Additionally, details on the specific commits and files associated with the pull request\'s most recent iteration have been saved to the Iteration_Details.txt file—these will serve as the focus for the current code review. The Iteration_Details.txt file also lists the full commit SHAs for the latest iteration: the Source Commit contains the pull request\'s changes (the new code), and the Target Commit is the tip of the base branch it will merge into (the old code). When using git to inspect the changes, always diff from base to PR branch using the merge-base form—git diff <target-commit>...<source-commit>—so that additions in the output represent code introduced by this pull request and deletions represent code it removes. If a Work_Item_Details.txt file exists in the working directory, it contains the full details of work items linked to this pull request, including their type, title, description, acceptance criteria, and repro steps. Use this information to better understand the intent and requirements behind the code changes being reviewed. If network conditions permit, you may pull more information directly from the Azure DevOps API using the PAT configured in the AZUREDEVOPSPAT environment variable if it would be useful.';
                const replacedOverview = 'The PR details, iteration information, and code diff are all provided inline at the end of this prompt. Use this information to understand the context and review the code changes. Do NOT attempt to read files from disk or run git commands to explore the repository—all relevant context is embedded below.';
                if (currentPrompt.includes(originalOverview)) {
                    currentPrompt = currentPrompt.replace(originalOverview, replacedOverview);
                } else {
                    tl.warning('Diff-only mode: expected Overview paragraph not found in the prompt template. The prompt template may have been edited; update the originalOverview string in index.ts to match.');
                }
            }

            // Build the inline context sections (appended for all prompt types)
            let contextSections = '\n\n# PR Details\n\n';
            if (fs.existsSync(prDetailsOutput)) {
                contextSections += fs.readFileSync(prDetailsOutput, 'utf8');
            }

            contextSections += '\n\n# Iteration Details\n\n';
            if (fs.existsSync(iterationDetailsOutput)) {
                contextSections += fs.readFileSync(iterationDetailsOutput, 'utf8');
            }

            const workItemDetailsFile = path.join(workingDirectory, 'Work_Item_Details.txt');
            if (fs.existsSync(workItemDetailsFile)) {
                contextSections += '\n\n# Work Item Details\n\n';
                contextSections += fs.readFileSync(workItemDetailsFile, 'utf8');
            }

            contextSections += '\n\n# Code Changes (git diff)\n\nThe following unified diff shows all code changes in this pull request:\n\n```diff\n' + diffContent + '\n```\n';

            // For non-raw prompts, append a strong override directive.
            // Raw prompt users provide their own diff-only instructions.
            if (!isRawPrompt) {
                const overrideBlock = '\n\n# IMPORTANT: Diff-Only Review Mode\n\n' +
                    'This is a diff-only review. All necessary context (PR details, iteration information, work items, and the full code diff) is provided in the sections above. You MUST:\n\n' +
                    '- Use ONLY the embedded diff and PR details for your review\n' +
                    '- DO NOT attempt to read files from disk or run git commands to explore the repository\n' +
                    '- DO NOT use shell commands for browsing or inspecting source files\n' +
                    '- Use only the pwsh-based comment scripts (Add-CopilotComment.ps1, Update-CopilotComment.ps1) to post your feedback\n\n' +
                    'Any earlier instructions in this prompt that conflict with the above are superseded by this section.\n';
                contextSections += overrideBlock;
            }

            // Append all context to the prompt
            currentPrompt += contextSections;

            // Write the assembled prompt
            const diffPromptPath = path.join(workingDirectory, '_diff_prompt.txt');
            fs.writeFileSync(diffPromptPath, currentPrompt, 'utf8');
            promptFilePath = diffPromptPath;
            diffOnlyActive = true;

            console.log('[Diff-Only Mode] All context embedded in prompt. Tool access will be restricted.');
        }

        // Copy the Add-AzureDevOpsPRComment.ps1 and Add-AzureDevOpsPRComment.ps1 script to the working directory
        // so Copilot can find and use them for posting PR comments
        const addCommentScriptSource = path.join(scriptsDir, 'Add-AzureDevOpsPRComment.ps1');
        const commentScriptSource = path.join(scriptsDir, 'Add-CopilotComment.ps1');
        const updateCommentScriptSource = path.join(scriptsDir, 'Update-CopilotComment.ps1');
        const deleteCommentScriptSource = path.join(scriptsDir, 'Delete-CopilotComment.ps1');
        const addCommentScriptDest = path.join(workingDirectory, 'Add-AzureDevOpsPRComment.ps1');
        const commentScriptDest = path.join(workingDirectory, 'Add-CopilotComment.ps1');
        const updateCommentScriptDest = path.join(workingDirectory, 'Update-CopilotComment.ps1');
        const deleteCommentScriptDest = path.join(workingDirectory, 'Delete-CopilotComment.ps1');
        fs.copyFileSync(addCommentScriptSource, addCommentScriptDest);
        console.log(`Copied Add-AzureDevOpsPRComment.ps1 to: ${addCommentScriptDest}`);
        fs.copyFileSync(commentScriptSource, commentScriptDest);
        console.log(`Copied Add-CopilotComment.ps1 to: ${commentScriptDest}`);
        fs.copyFileSync(updateCommentScriptSource, updateCommentScriptDest);
        console.log(`Copied Update-CopilotComment.ps1 to: ${updateCommentScriptDest}`);
        fs.copyFileSync(deleteCommentScriptSource, deleteCommentScriptDest);
        console.log(`Copied Delete-CopilotComment.ps1 to: ${deleteCommentScriptDest}`);
        
        // Publish prompt artifacts for debugging if requested
        if (publishPromptArtifacts) {
            console.log('\n[Artifacts] Publishing prompt and context files...');
            const artifactName = 'CopilotCodeReview';
            const artifactFiles = [
                prDetailsOutput,
                iterationDetailsOutput,
                path.join(workingDirectory, 'Iteration_Id.txt'),
                path.join(workingDirectory, 'Source_Commit.txt'),
                path.join(workingDirectory, 'Target_Commit.txt'),
                path.join(workingDirectory, 'Work_Item_Ids.txt'),
                path.join(workingDirectory, 'Work_Item_Details.txt'),
                promptFilePath,
            ];

            for (const artifactFile of artifactFiles) {
                if (artifactFile && fs.existsSync(artifactFile)) {
                    tl.command('artifact.upload', {
                        containerfolder: artifactName,
                        artifactname: artifactName
                    }, artifactFile);
                    console.log(`  Uploaded: ${path.basename(artifactFile)}`);
                }
            }
        }

        // Run CLI agent with timeout
        const timeoutMs = timeoutMinutes * 60 * 1000;
        if (useClaudeCode) {
            await runClaudeCodeCli(promptFilePath, model, workingDirectory, timeoutMs, maxTurns, maxBudget, diffOnlyActive);
        } else {
            await runCopilotCli(promptFilePath, model, workingDirectory, timeoutMs, diffOnlyActive);
        }

        console.log('\n' + '='.repeat(60));
        console.log(`${agentName} Code Review completed successfully!`);
        console.log('='.repeat(60));

        tl.setResult(tl.TaskResult.Succeeded, `${agentName} code review completed.`);
    } catch (err: unknown) {
        const errorMessage = err instanceof Error ? err.message : String(err);
        tl.setResult(tl.TaskResult.Failed, `Task failed: ${errorMessage}`);
    }
}

async function checkCopilotCli(): Promise<boolean> {
    try {
        const result = child_process.spawnSync('copilot', ['--version'], {
            encoding: 'utf8',
            shell: true
        });
        return result.status === 0;
    } catch {
        return false;
    }
}

/**
 * Runs a single installer command to completion, resolving on exit code 0.
 */
function runInstallerOnce(command: string, args: string[]): Promise<void> {
    return new Promise((resolve, reject) => {
        const installProcess = child_process.spawn(command, args, {
            shell: true,
            stdio: 'inherit'
        });

        installProcess.on('close', (code: number | null) => {
            if (code === 0) {
                resolve();
            } else {
                reject(new Error(`Exit code: ${code}`));
            }
        });

        installProcess.on('error', (err: Error) => {
            reject(err);
        });
    });
}

/**
 * Runs an installer command, retrying with exponential backoff on failure.
 * CLI installs regularly hit transient CDN errors — e.g. GitHub returning 504
 * on release assets (issue #57) — so failed attempts are usually worth
 * retrying before failing the whole task.
 */
async function runInstallerWithRetry(
    description: string,
    command: string,
    args: string[],
    maxAttempts: number = 3
): Promise<void> {
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            await runInstallerOnce(command, args);
            return;
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            if (attempt === maxAttempts) {
                throw new Error(`${description} failed after ${maxAttempts} attempts. Last error: ${message}`);
            }
            const delaySeconds = Math.pow(2, attempt);
            console.log(`${description} failed (attempt ${attempt} of ${maxAttempts}): ${message}`);
            console.log(`Retrying in ${delaySeconds} seconds...`);
            await new Promise(resolve => setTimeout(resolve, delaySeconds * 1000));
        }
    }
}

async function installCopilotCli(): Promise<void> {
    let command: string;
    let args: string[];

    if (isWindows()) {
        console.log('Installing GitHub Copilot CLI via winget...');
        command = 'winget';
        args = ['install', 'GitHub.Copilot', '--silent', '--accept-package-agreements', '--accept-source-agreements'];
    } else {
        console.log('Installing GitHub Copilot CLI via official install script...');
        // Use the official GitHub install script which downloads a pre-built binary
        // The script installs to $HOME/.local/bin by default for non-root users
        // Pass the full command as a single string when using shell: true.
        // curl's --retry covers transient failures fetching the install script
        // itself; the outer retry loop covers failures inside the piped script,
        // e.g. 504s downloading the release tarball (issue #57).
        command = 'curl -fsSL --retry 3 --retry-all-errors https://gh.io/copilot-install | bash';
        args = [];
    }

    await runInstallerWithRetry('GitHub Copilot CLI installation', command, args);

    console.log('GitHub Copilot CLI installed successfully.');
    // On Linux, add the install location to PATH for the current process
    if (!isWindows()) {
        const homeDir = process.env['HOME'] || '';
        const localBin = path.join(homeDir, '.local', 'bin');
        process.env['PATH'] = `${localBin}:${process.env['PATH']}`;
        console.log(`Added ${localBin} to PATH.`);
    }
}

async function runPowerShellScript(scriptPath: string, args: string[]): Promise<void> {
    return new Promise((resolve, reject) => {
        const command = `pwsh -NoProfile -File "${scriptPath}" ${args.join(' ')}`;
        const envVars = { ...process.env };
        
        const psProcess = child_process.spawn(command, [], {
            shell: true,
            stdio: 'inherit',
            env: envVars
        });

        psProcess.on('close', (code) => {
            if (code === 0) {
                resolve();
            } else {
                reject(new Error(`PowerShell script failed with exit code: ${code}`));
            }
        });

        psProcess.on('error', (err) => {
            reject(new Error(`Failed to run PowerShell script: ${err.message}`));
        });
    });
}

async function runCopilotCli(promptFilePath: string, model: string | undefined, workingDirectory: string, timeoutMs: number, diffOnlyActive: boolean): Promise<void> {
    return new Promise((resolve, reject) => {
        // Build PowerShell command that reads prompt file and passes content to copilot CLI
        let copilotFlags = `--allow-all-paths --allow-all-tools --deny-tool 'shell(git push)'`;
        if (model) {
            copilotFlags += ` --model ${model}`;
        }

        // In diff-only mode the prompt contains the embedded diff and is too noisy to print.
        // Users can inspect the full prompt via publishPromptArtifacts.
        const printPrompt = diffOnlyActive
            ? `Write-Host '[Diff-only mode] Prompt suppressed; enable publishPromptArtifacts to inspect the full prompt.';`
            : `Write-Host ========== START PROMPT ==========; Write-Host ($prompt -replace '(?m)^##', ' ##'); Write-Host ========== END PROMPT ==========;`;
        const envRefresh = `$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User");`
        // Pipe the prompt via stdin — avoids Windows command-line length limits and
        // eliminates double-quote escaping issues in prompt content.
        const copilotInvocation = `$prompt | copilot ${copilotFlags}`;
        const sanitizedCopilotCmd = `$ErrorActionPreference = 'Stop'; try { ${copilotInvocation} | ForEach-Object { Write-Host ($_ -replace '(?m)^##', ' ##') }; if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE } } catch { Write-Host "Error running Copilot CLI: $_"; exit 1 }`;
        const psCommand = `${envRefresh} $prompt = Get-Content -Path '${promptFilePath}' -Raw; ${printPrompt} ${sanitizedCopilotCmd}`;
        console.log(`Running Powershell: ${psCommand}`);
        
        const envVars = { ...process.env };
        
        const copilotProcess = child_process.spawn(
            'pwsh',
            ['-NoProfile', '-Command', psCommand],
            {
                shell: false,
                stdio: 'inherit',
                cwd: workingDirectory,
                env: envVars
            }
        );

        // Set up timeout
        const timeoutId = setTimeout(() => {
            console.log(`\nTimeout reached (${timeoutMs / 60000} minutes). Terminating Copilot process...`);
            copilotProcess.kill('SIGTERM');
            reject(new Error(`Copilot review timed out after ${timeoutMs / 60000} minutes`));
        }, timeoutMs);

        copilotProcess.on('close', (code) => {
            clearTimeout(timeoutId);
            if (code === 0) {
                resolve();
            } else {
                reject(new Error(`Copilot CLI exited with code: ${code}`));
            }
        });

        copilotProcess.on('error', (err) => {
            clearTimeout(timeoutId);
            reject(new Error(`Failed to run Copilot CLI: ${err.message}`));
        });
    });
}

async function checkClaudeCodeCli(): Promise<boolean> {
    try {
        const result = child_process.spawnSync('claude', ['--version'], {
            encoding: 'utf8',
            shell: true
        });
        return result.status === 0;
    } catch {
        return false;
    }
}

async function installClaudeCodeCli(): Promise<void> {
    console.log('Installing Claude Code CLI via npm...');
    await runInstallerWithRetry('Claude Code CLI installation', 'npm', ['install', '-g', '@anthropic-ai/claude-code']);

    console.log('Claude Code CLI installed successfully.');
    // Ensure npm global bin is on PATH for the current process
    const npmBinResult = child_process.spawnSync('npm', ['bin', '-g'], {
        encoding: 'utf8',
        shell: true
    });
    if (npmBinResult.status === 0 && npmBinResult.stdout.trim()) {
        const npmGlobalBin = npmBinResult.stdout.trim();
        const pathSep = isWindows() ? ';' : ':';
        process.env['PATH'] = `${npmGlobalBin}${pathSep}${process.env['PATH']}`;
        console.log(`Added ${npmGlobalBin} to PATH.`);
    }
}

async function runClaudeCodeCli(
    promptFilePath: string,
    model: string | undefined,
    workingDirectory: string,
    timeoutMs: number,
    maxTurns: string | undefined,
    maxBudget: string | undefined,
    diffOnlyActive: boolean
): Promise<void> {
    return new Promise((resolve, reject) => {
        // Build Claude Code CLI flags
        let claudeFlags = `--dangerously-skip-permissions`;
        claudeFlags += ` --output-format stream-json --verbose --include-partial-messages`;
        if (diffOnlyActive) {
            claudeFlags += ` --allowedTools "Bash(pwsh *)" "Write"`;
        } else {
            claudeFlags += ` --allowedTools "Bash" "Read" "Write" "Edit" "Glob" "Grep"`;
        }
        claudeFlags += ` --disallowedTools "Bash(git push *)"`;

        if (model) {
            claudeFlags += ` --model ${model}`;
        }
        if (maxTurns) {
            claudeFlags += ` --max-turns ${maxTurns}`;
        }
        if (maxBudget) {
            claudeFlags += ` --max-budget-usd ${maxBudget}`;
        }

        // When diffOnlyActive, pipe the prompt via stdin to avoid Windows command-line
        // length limits. Claude Code reads from stdin when content is piped to `claude -p`.
        const claudeInvocation = diffOnlyActive
            ? `$prompt | claude -p ${claudeFlags}`
            : `claude -p "$prompt" ${claudeFlags}`;

        // PowerShell pipeline: stream JSON from Claude Code, extract text deltas and tool calls in real time.
        // Tool use events arrive as: content_block_start (tool name) → input_json_delta fragments → content_block_stop.
        // We accumulate input fragments with $script:-scoped state and print the tool name + input on block stop.
        // $ErrorActionPreference = 'Stop' ensures process startup failures are catchable.
        const streamParser = [
            `$ErrorActionPreference = 'Stop';`,
            `$script:toolNames = @{};`,
            `$script:toolInputs = @{};`,
            `try {`,
            `${claudeInvocation} | ForEach-Object {`,
            `  try {`,
            `    $ev = ($_ | ConvertFrom-Json -ErrorAction Stop);`,
            `    if ($ev.type -ne 'stream_event') { return }`,
            `    $se = $ev.event;`,
            `    if ($se.type -eq 'content_block_delta' -and $se.delta.type -eq 'text_delta') {`,
            `      Write-Host ($se.delta.text -replace '(?m)^##', ' ##')`,
            `    }`,
            `    elseif ($se.type -eq 'content_block_start' -and $se.content_block.type -eq 'tool_use') {`,
            `      $idx = [string]$se.index;`,
            `      $script:toolNames[$idx] = $se.content_block.name;`,
            `      $script:toolInputs[$idx] = ''`,
            `    }`,
            `    elseif ($se.type -eq 'content_block_delta' -and $se.delta.type -eq 'input_json_delta') {`,
            `      $idx = [string]$se.index;`,
            `      $script:toolInputs[$idx] += $se.delta.partial_json`,
            `    }`,
            `    elseif ($se.type -eq 'content_block_stop' -and $script:toolNames.ContainsKey([string]$se.index)) {`,
            `      $idx = [string]$se.index;`,
            `      $name = $script:toolNames[$idx];`,
            `      $raw = $script:toolInputs[$idx];`,
            `      try { $inp = ($raw | ConvertFrom-Json -ErrorAction Stop); $display = ($inp | ConvertTo-Json -Compress) } catch { $display = $raw };`,
            `      Write-Host '';`,
            `      Write-Host ("    [$name] $display" -replace '(?m)^##', ' ##');`,
            `      $script:toolNames.Remove($idx);`,
            `      $script:toolInputs.Remove($idx)`,
            `    }`,
            `  } catch { Write-Host $_ }`,
            `};`,
            `Write-Host '';`,
            `if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }`,
            `} catch { Write-Host "Error running Claude Code CLI: $_"; exit 1 }`
        ].join(' ');

        // In diff-only mode the prompt contains the embedded diff and is too noisy to print.
        // Users can inspect the full prompt via publishPromptArtifacts.
        const printPrompt = diffOnlyActive
            ? `Write-Host '[Diff-only mode] Prompt suppressed; enable publishPromptArtifacts to inspect the full prompt.';`
            : `Write-Host ========== START PROMPT ==========; Write-Host ($prompt -replace '(?m)^##', ' ##'); Write-Host ========== END PROMPT ==========;`;
        const envRefresh = isWindows()
            ? `$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User");`
            : '';
        const psCommand = `${envRefresh} $prompt = Get-Content -Path '${promptFilePath}' -Raw; ${printPrompt} ${streamParser}`;
        console.log(`Running PowerShell: ${psCommand}`);

        const envVars = { ...process.env };

        const claudeProcess = child_process.spawn(
            'pwsh',
            ['-NoProfile', '-Command', psCommand],
            {
                shell: false,
                stdio: 'inherit',
                cwd: workingDirectory,
                env: envVars
            }
        );

        // Set up timeout
        const timeoutId = setTimeout(() => {
            console.log(`\nTimeout reached (${timeoutMs / 60000} minutes). Terminating Claude Code process...`);
            claudeProcess.kill('SIGTERM');
            reject(new Error(`Claude Code review timed out after ${timeoutMs / 60000} minutes`));
        }, timeoutMs);

        claudeProcess.on('close', (code) => {
            clearTimeout(timeoutId);
            if (code === 0) {
                resolve();
            } else {
                reject(new Error(`Claude Code CLI exited with code: ${code}`));
            }
        });

        claudeProcess.on('error', (err) => {
            clearTimeout(timeoutId);
            reject(new Error(`Failed to run Claude Code CLI: ${err.message}`));
        });
    });
}

// Run only when executed as the task entry point (node index.js). Importing
// this module (e.g. from tests) must not kick off a task run.
if (require.main === module) {
    run();
}

// Exported for tests
export { runInstallerWithRetry, normalizeGitHubHost, buildReviewDirectives };
