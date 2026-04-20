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
        const anthropicApiKey = tl.getInput('anthropicApiKey');
        const maxTurns = tl.getInput('maxTurns');
        const maxBudget = tl.getInput('maxBudget');
        const copilotProviderBaseUrl = tl.getInput('copilotProviderBaseUrl');
        const copilotProviderType = tl.getInput('copilotProviderType');
        const copilotProviderApiKey = tl.getInput('copilotProviderApiKey');

        // Validate agent-specific auth
        if (useClaudeCode) {
            if (!anthropicApiKey) {
                tl.setResult(tl.TaskResult.Failed,
                    'Anthropic API Key is required when using Claude Code CLI. Please provide the anthropicApiKey input.');
                return;
            }
        } else if (useCustomModelProvider) {
            // copilot CLI with custom model provider
            if (!copilotProviderBaseUrl) {
                tl.setResult(tl.TaskResult.Failed,
                    'Base URL of model provider endpoint is required when using a custom model provider.');
                return;
            }
            if (!copilotProviderType) {
                tl.setResult(tl.TaskResult.Failed,
                    'Type of model provider is required when using a custom model provider.');
                return;
            }
            if (!copilotProviderApiKey) {
                tl.setResult(tl.TaskResult.Failed,
                    'API key of model provider is required when using a custom model provider.');
                return;
            }
        } else {
            // copilot CLI with GitHub-hosted models
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
        const model = tl.getInput('model');
        const promptFile = tl.getInput('promptFile');
        const prompt = tl.getInput('prompt');
        const promptRaw = tl.getInput('promptRaw');
        const promptFileRaw = tl.getInput('promptFileRaw');
        const includeWorkItems = tl.getBoolInput('includeWorkItems', false);

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
        console.log(`Collection URI: ${resolvedCollectionUri}`);
        console.log(`Project: ${project}`);
        console.log(`Repository: ${repository}`);
        console.log(`Pull Request ID: ${pullRequestId}`);
        console.log(`Timeout: ${timeoutMinutes} minutes`);
        if (model) {
            console.log(`Model: ${model}`);
        }
        console.log('='.repeat(60));

        // Set environment variables for PowerShell scripts
        if (useClaudeCode) {
            process.env['ANTHROPIC_API_KEY'] = anthropicApiKey!;
        } else if (useCustomModelProvider) {
            process.env['COPILOT_PROVIDER_BASE_URL'] = copilotProviderBaseUrl!;
            process.env['COPILOT_PROVIDER_TYPE'] = copilotProviderType!;
            process.env['COPILOT_PROVIDER_API_KEY'] = copilotProviderApiKey!;
        } else {
            process.env['GH_TOKEN'] = githubPat!;
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

        // Helper to check if filePath inputs are actually set (filePath inputs return working dir when empty)
        const isPromptFileSet = promptFile &&
            fs.existsSync(promptFile) &&
            fs.statSync(promptFile).isFile();
        const isPromptFileRawSet = promptFileRaw &&
            fs.existsSync(promptFileRaw) &&
            fs.statSync(promptFileRaw).isFile();

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
            console.log('\nRAW PROMPT:\n' + promptRaw + '\n\n');
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
            console.log('\nRAW PROMPT:\n' + fileContent + '\n\n');
        } else if (prompt) {
            // Direct prompt input: merge with template
            console.log('Using custom prompt from input.');
            if (prompt.includes('"')) {
                tl.setResult(tl.TaskResult.Failed, 'Custom prompts cannot include double quotes ("). Please remove any double quotes from your prompt input.');
                return;
            }
            customPromptText = prompt;
        } else if (isPromptFileSet) {
            // Read from prompt file: merge with template
            console.log(`Using custom prompt from file: ${promptFile}`);
            const fileContent = fs.readFileSync(promptFile!, 'utf8').trim();
            if (!fileContent) {
                tl.setResult(tl.TaskResult.Failed, `Prompt file is empty: ${promptFile}`);
                return;
            }
            if (fileContent.includes('"')) {
                tl.setResult(tl.TaskResult.Failed, `Custom prompts cannot include double quotes ("). Please remove any double quotes from the prompt file: ${promptFile}`);
                return;
            }
            customPromptText = fileContent;
        }

        if (customPromptText) {
            // Use custom prompt template with placeholder replacement
            const customPromptTemplate = path.join(scriptsDir, 'prompt-custom.txt');
            const templateContent = fs.readFileSync(customPromptTemplate, 'utf8');
            const mergedPrompt = templateContent.replace('%CUSTOMPROMPT%', customPromptText);
            console.log('\nCUSTOM PROMPT:\n' + mergedPrompt + '\n\n');

            // Write merged prompt to a temp file in the working directory
            promptFilePath = path.join(workingDirectory, '_copilot_prompt.txt');
            fs.writeFileSync(promptFilePath, mergedPrompt, 'utf8');
            console.log('Custom prompt merged with instruction template.');
        } else if (!promptRaw && !isPromptFileRawSet) {
            // Use default prompt file bundled with the task
            promptFilePath = path.join(scriptsDir, 'prompt.txt');
            console.log('Using default prompt.');
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
        
        // Run CLI agent with timeout
        const timeoutMs = timeoutMinutes * 60 * 1000;
        if (useClaudeCode) {
            await runClaudeCodeCli(promptFilePath, model, workingDirectory, timeoutMs, maxTurns, maxBudget);
        } else {
            await runCopilotCli(promptFilePath, model, workingDirectory, timeoutMs);
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

async function installCopilotCli(): Promise<void> {
    return new Promise((resolve, reject) => {
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
            // Pass the full command as a single string when using shell: true
            command = 'curl -fsSL https://gh.io/copilot-install | bash';
            args = [];
        }
        
        const installProcess = child_process.spawn(
            command,
            args,
            {
                shell: true,
                stdio: 'inherit'
            }
        );

        installProcess.on('close', (code: number | null) => {
            if (code === 0) {
                console.log('GitHub Copilot CLI installed successfully.');
                // On Linux, add the install location to PATH for the current process
                if (!isWindows()) {
                    const homeDir = process.env['HOME'] || '';
                    const localBin = path.join(homeDir, '.local', 'bin');
                    process.env['PATH'] = `${localBin}:${process.env['PATH']}`;
                    console.log(`Added ${localBin} to PATH.`);
                }
                resolve();
            } else {
                reject(new Error(`Failed to install GitHub Copilot CLI. Exit code: ${code}`));
            }
        });

        installProcess.on('error', (err: Error) => {
            reject(new Error(`Failed to install GitHub Copilot CLI: ${err.message}`));
        });
    });
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

async function runCopilotCli(promptFilePath: string, model: string | undefined, workingDirectory: string, timeoutMs: number): Promise<void> {
    return new Promise((resolve, reject) => {
        // Build PowerShell command that reads prompt file and passes content to copilot CLI
        // This mirrors the original implementation: $prompt = Get-Content -Path "prompt.txt" -Raw; copilot -p $prompt ...
        let copilotCmd = `copilot -p "$prompt" --allow-all-paths --allow-all-tools --deny-tool 'shell(git push)'`;
        if (model) {
            copilotCmd += ` --model ${model}`;
        }
        
        const printPrompt = `Write-Host ========== START PROMPT ==========; Write-Host $prompt; Write-Host ========== END PROMPT ==========;`;
        const envRefresh = `$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User");`
        const psCommand = `${envRefresh} $prompt = Get-Content -Path '${promptFilePath}' -Raw; ${printPrompt} ${copilotCmd}`;
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
    return new Promise((resolve, reject) => {
        console.log('Installing Claude Code CLI via npm...');
        const installProcess = child_process.spawn(
            'npm',
            ['install', '-g', '@anthropic-ai/claude-code'],
            {
                shell: true,
                stdio: 'inherit'
            }
        );

        installProcess.on('close', (code: number | null) => {
            if (code === 0) {
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
                resolve();
            } else {
                reject(new Error(`Failed to install Claude Code CLI. Exit code: ${code}`));
            }
        });

        installProcess.on('error', (err: Error) => {
            reject(new Error(`Failed to install Claude Code CLI: ${err.message}`));
        });
    });
}

async function runClaudeCodeCli(
    promptFilePath: string,
    model: string | undefined,
    workingDirectory: string,
    timeoutMs: number,
    maxTurns: string | undefined,
    maxBudget: string | undefined
): Promise<void> {
    return new Promise((resolve, reject) => {
        // Build Claude Code CLI command for headless CI/CD operation
        // Use stream-json output with PowerShell parsing for real-time log streaming
        let claudeCmd = `claude -p "$prompt" --dangerously-skip-permissions`;
        claudeCmd += ` --output-format stream-json --verbose --include-partial-messages`;
        claudeCmd += ` --allowedTools "Bash" "Read" "Write" "Edit" "Glob" "Grep"`;
        claudeCmd += ` --disallowedTools "Bash(git push *)"`;

        if (model) {
            claudeCmd += ` --model ${model}`;
        }
        if (maxTurns) {
            claudeCmd += ` --max-turns ${maxTurns}`;
        }
        if (maxBudget) {
            claudeCmd += ` --max-budget-usd ${maxBudget}`;
        }

        // PowerShell pipeline: stream JSON from Claude Code, extract text deltas and tool calls in real time.
        // Tool use events arrive as: content_block_start (tool name) → input_json_delta fragments → content_block_stop.
        // We accumulate input fragments with $script:-scoped state and print the tool name + input on block stop.
        const streamParser = [
            `$script:toolNames = @{};`,
            `$script:toolInputs = @{};`,
            `${claudeCmd} | ForEach-Object {`,
            `  try {`,
            `    $ev = ($_ | ConvertFrom-Json -ErrorAction Stop);`,
            `    if ($ev.type -ne 'stream_event') { return }`,
            `    $se = $ev.event;`,
            `    if ($se.type -eq 'content_block_delta' -and $se.delta.type -eq 'text_delta') {`,
            `      Write-Host $se.delta.text`,
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
            `      Write-Host "    [$name] $display";`,
            `      $script:toolNames.Remove($idx);`,
            `      $script:toolInputs.Remove($idx)`,
            `    }`,
            `  } catch { Write-Host $_ }`,
            `};`,
            `Write-Host '';`,
            `if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }`
        ].join(' ');

        const printPrompt = `Write-Host ========== START PROMPT ==========; Write-Host $prompt; Write-Host ========== END PROMPT ==========;`;
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

run();
