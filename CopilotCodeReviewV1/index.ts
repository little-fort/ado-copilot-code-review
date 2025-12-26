import * as tl from 'azure-pipelines-task-lib/task';
import * as path from 'path';
import * as fs from 'fs';
import * as child_process from 'child_process';

async function run(): Promise<void> {
    try {
        // Get required inputs
        const githubPat = tl.getInputRequired('githubPat');
        const azureDevOpsPat = tl.getInputRequired('azureDevOpsPat');
        const organization = tl.getInputRequired('organization');
        const project = tl.getInputRequired('project');
        const repository = tl.getInputRequired('repository');

        // Get optional inputs
        let pullRequestId = tl.getInput('pullRequestId');
        const timeoutMinutes = parseInt(tl.getInput('timeout') || '15', 10);
        const model = tl.getInput('model');
        const promptFile = tl.getInput('promptFile');
        const prompt = tl.getInput('prompt');

        // If PR ID not provided, try to get from pipeline variable
        if (!pullRequestId) {
            pullRequestId = tl.getVariable('System.PullRequest.PullRequestId');
        }

        if (!pullRequestId) {
            tl.setResult(tl.TaskResult.Failed, 'Pull Request ID is required. Either provide it as an input or run this task as part of a PR validation build.');
            return;
        }

        console.log('='.repeat(60));
        console.log('Copilot Code Review Task');
        console.log('='.repeat(60));
        console.log(`Organization: ${organization}`);
        console.log(`Project: ${project}`);
        console.log(`Repository: ${repository}`);
        console.log(`Pull Request ID: ${pullRequestId}`);
        console.log(`Timeout: ${timeoutMinutes} minutes`);
        if (model) {
            console.log(`Model: ${model}`);
        }
        console.log('='.repeat(60));

        // Set environment variables for PowerShell scripts
        process.env['GH_TOKEN'] = githubPat;
        process.env['AZUREDEVOPSPAT'] = azureDevOpsPat;
        process.env['ORGANIZATION'] = organization;
        process.env['PROJECT'] = project;
        process.env['REPOSITORY'] = repository;
        process.env['PRID'] = pullRequestId;

        const scriptsDir = path.join(__dirname, 'scripts');
        const workingDirectory = tl.getVariable('System.DefaultWorkingDirectory') || process.cwd();

        // Step 1: Install GitHub Copilot CLI if not present
        console.log('\n[Step 1/4] Checking GitHub Copilot CLI installation...');
        const copilotInstalled = await checkCopilotCli();
        if (!copilotInstalled) {
            console.log('GitHub Copilot CLI not found. Installing...');
            await installCopilotCli();
        } else {
            console.log('GitHub Copilot CLI is already installed.');
        }

        // Step 2: Fetch PR details
        console.log('\n[Step 2/4] Fetching pull request details...');
        const prDetailsScript = path.join(scriptsDir, 'Get-AzureDevOpsPR.ps1');
        const prDetailsOutput = path.join(workingDirectory, 'PR_Details.txt');
        
        await runPowerShellScript(prDetailsScript, [
            `-PAT "${azureDevOpsPat}"`,
            `-Organization "${organization}"`,
            `-Project "${project}"`,
            `-Repository "${repository}"`,
            `-Id ${pullRequestId}`,
            `-OutputFile "${prDetailsOutput}"`
        ]);
        console.log(`PR details saved to: ${prDetailsOutput}`);

        // Step 3: Fetch PR changes (iteration details)
        console.log('\n[Step 3/4] Fetching pull request changes...');
        const prChangesScript = path.join(scriptsDir, 'Get-AzureDevOpsPRChanges.ps1');
        const iterationDetailsOutput = path.join(workingDirectory, 'Iteration_Details.txt');
        
        await runPowerShellScript(prChangesScript, [
            `-PAT "${azureDevOpsPat}"`,
            `-Organization "${organization}"`,
            `-Project "${project}"`,
            `-Repository "${repository}"`,
            `-Id ${pullRequestId}`,
            `-OutputFile "${iterationDetailsOutput}"`
        ]);
        console.log(`Iteration details saved to: ${iterationDetailsOutput}`);

        // Step 4: Run Copilot CLI for code review
        console.log('\n[Step 4/4] Running Copilot code review...');
        
        // Determine the prompt to use
        let copilotPrompt: string;
        let customPromptText: string | null = null;
        
        // Helper to check if promptFile is actually set (filePath inputs return working dir when empty)
        const isPromptFileSet = promptFile && 
            fs.existsSync(promptFile) && 
            fs.statSync(promptFile).isFile();
        
        if (prompt) {
            // Direct prompt input takes precedence
            console.log('Using custom prompt from input.');
            customPromptText = prompt;
        } else if (isPromptFileSet) {
            // Read from prompt file
            console.log(`Using custom prompt from file: ${promptFile}`);
            const fileContent = fs.readFileSync(promptFile, 'utf8').trim();
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
            copilotPrompt = templateContent.replace('%CUSTOMPROMPT%', customPromptText);
            console.log('Custom prompt merged with instruction template.');
        } else {
            // Use default prompt file bundled with the task
            const defaultPromptFile = path.join(scriptsDir, 'prompt.txt');
            console.log('Using default prompt.');
            copilotPrompt = fs.readFileSync(defaultPromptFile, 'utf8');
        }

        // Build copilot command arguments
        const copilotArgs: string[] = [
            '-p', `"${copilotPrompt.replace(/"/g, '\\"')}"`,
            '--allow-all-paths',
            '--allow-all-tools'
        ];

        // Add model if specified
        if (model) {
            copilotArgs.push('--model', model);
        }

        // Run Copilot CLI with timeout
        const timeoutMs = timeoutMinutes * 60 * 1000;
        await runCopilotCli(copilotArgs, workingDirectory, timeoutMs);

        console.log('\n' + '='.repeat(60));
        console.log('Copilot Code Review completed successfully!');
        console.log('='.repeat(60));

        tl.setResult(tl.TaskResult.Succeeded, 'Copilot code review completed.');
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
        console.log('Installing GitHub Copilot CLI via winget...');
        
        const installProcess = child_process.spawn(
            'winget',
            ['install', 'GitHub.Copilot', '--silent', '--accept-package-agreements', '--accept-source-agreements'],
            {
                shell: true,
                stdio: 'inherit'
            }
        );

        installProcess.on('close', (code: number | null) => {
            if (code === 0) {
                console.log('GitHub Copilot CLI installed successfully.');
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
        const command = `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "${scriptPath}" ${args.join(' ')}`;
        
        const psProcess = child_process.spawn(command, [], {
            shell: true,
            stdio: 'inherit',
            env: { ...process.env }
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

async function runCopilotCli(args: string[], workingDirectory: string, timeoutMs: number): Promise<void> {
    return new Promise((resolve, reject) => {
        const command = `copilot ${args.join(' ')}`;
        console.log(`Running: copilot [prompt] --allow-all-paths --allow-all-tools${args.includes('--model') ? ' --model ' + args[args.indexOf('--model') + 1] : ''}`);
        
        const copilotProcess = child_process.spawn(command, [], {
            shell: true,
            stdio: 'inherit',
            cwd: workingDirectory,
            env: { ...process.env }
        });

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

run();
