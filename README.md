# Copilot Code Review for Azure DevOps

[![Azure DevOps Marketplace](https://img.shields.io/badge/Azure%20DevOps-Marketplace-blue)](https://marketplace.visualstudio.com/items?itemName=LittleFortSoftware.copilot-code-review)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Automated pull request code reviews powered by the official GitHub Copilot CLI. Get AI-driven feedback on your PRs directly in Azure DevOps.

## Overview

This Azure DevOps extension provides a pipeline task that automatically reviews pull request code changes using GitHub Copilot. When triggered, the task:

1. Fetches pull request details and changed files from Azure DevOps
2. Invokes GitHub Copilot CLI to analyze the changes
3. Posts review comments directly to the pull request

This brings GitHub Copilot's code review capabilities to Azure DevOps, helping teams improve code quality through AI-assisted reviews.

## Prerequisites

Before using this extension, ensure you have:

- **GitHub Copilot Subscription**: An active GitHub Copilot subscription (Individual, Business, or Enterprise)
- **GitHub Personal Access Token**: A PAT with Copilot access permissions
- **Azure DevOps Personal Access Token**: A PAT with permissions to:
  - Read pull requests
  - Write pull request comments
  - Read code
- **Windows Agent**: This extension currently only supports Windows-based Azure DevOps agents. Compatible with both MS-hosted and self-hosted agents.

## Installation

1. Install the extension from the [Azure DevOps Marketplace](https://marketplace.visualstudio.com/items?itemName=LittleFortSoftware.copilot-code-review)
2. Navigate to your Azure DevOps organization settings
3. Go to **Extensions** and verify the extension is installed

## Usage

### Basic Usage

Add the task to your pipeline YAML:

```yaml
pr:
- main
- develop

pool:
  vmImage: 'windows-latest'

steps:
- checkout: self
  fetchDepth: 0

- task: CopilotCodeReview@1
  displayName: 'Copilot Code Review'
  inputs:
    githubPat: '$(GITHUB_PAT)'
    azureDevOpsPat: '$(AZURE_DEVOPS_PAT)'
    organization: 'your-org'
    project: 'your-project'
    repository: 'your-repo'
```

### With Custom Prompt

If the included prompt is not to your liking, you can customize the review prompt to focus on aspects tailored to your needs:

```yaml
- task: CopilotCodeReview@1
  displayName: 'Copilot Code Review'
  inputs:
    githubPat: '$(GITHUB_PAT)'
    azureDevOpsPat: '$(AZURE_DEVOPS_PAT)'
    organization: 'your-org'
    project: 'your-project'
    repository: 'your-repo'
    prompt: |
      Review this code focusing on:
      - Security vulnerabilities
      - SQL injection risks
      - Authentication/authorization issues
      Post any findings as PR comments.
```

For longer prompts, create a .txt file in your repository and pass the file path as a task input:

```yaml
- task: CopilotCodeReview@1
  displayName: 'Copilot Code Review'
  inputs:
    githubPat: '$(GITHUB_PAT)'
    azureDevOpsPat: '$(AZURE_DEVOPS_PAT)'
    organization: 'your-org'
    project: 'your-project'
    repository: 'your-repo'
    promptFile: '$(Build.SourcesDirectory)/.copilot/review-prompt.txt'
```

### Manual Trigger for Specific PR

If you don't want to setup an automatic trigger, you can instead manually pass in a pull request ID to run reviews on demand:

```yaml
parameters:
  - name: pullRequestId
    displayName: 'Pull Request ID'
    type: string
    default: ''

trigger: none

pool:
  vmImage: 'windows-latest'

steps:
  - checkout: self
    fetchDepth: 0

  - task: CopilotCodeReview@1
    displayName: 'Copilot Code Review'
    inputs:
      githubPat: '$(GITHUB_PAT)'
      azureDevOpsPat: '$(AZURE_DEVOPS_PAT)'
      organization: 'your-org'
      project: 'your-project'
      repository: 'your-repo'
      pullRequestId: '${{ parameters.pullRequestId }}'
```

## Input Reference

| Input | Required | Default | Description |
|-------|----------|---------|-------------|
| `githubPat` | Yes | - | GitHub Personal Access Token with Copilot access |
| `azureDevOpsPat` | Yes | - | Azure DevOps PAT for API access |
| `organization` | Yes | - | Azure DevOps organization name |
| `project` | Yes | - | Azure DevOps project name |
| `repository` | Yes | - | Repository name |
| `pullRequestId` | No | `$(System.PullRequest.PullRequestId)` | PR ID (auto-detected in PR builds) |
| `timeout` | No | `15` | Timeout in minutes |
| `model` | No | - | Preferred Copilot model to use (see valid options below) |
| `promptFile` | No | - | Path to custom prompt file |
| `prompt` | No | - | Inline custom prompt (overrides `promptFile`) |

### Copilot Models

As of December 2025, here are the model options supported by the GitHub Copilot CLI:

- `claude-sonnet-4.5` (default)
- `claude-haiku-4.5`
- `claude-opus-4.5`
- `claude-sonnet-4`
- `gpt-5.2`
- `gpt-5.1-codex-max`
- `gpt-5.1-codex`
- `gpt-5.1-codex-mini`
- `gpt-5.1`
- `gpt-5`
- `gpt-5-mini`
- `gpt-4.1`
- `gemini-3-pro-preview`

## Setting Up Tokens

### GitHub Personal Access Token

1. Go to [GitHub Settings > Developer Settings > Personal Access Tokens](https://github.com/settings/tokens)
2. Generate a new **Fine-grained** token with the following options:
   - Repository access: Public
   - Permission: Copilot Requests
3. Store the token as a secret variable in your Azure DevOps pipeline

### Azure DevOps Personal Access Token

1. Go to your Azure DevOps organization
2. Click on **User Settings** > **Personal Access Tokens**
3. Click on **New Token** and then **Show All Scopes**
4. Create a new token with the following scopes:
   - **Code**: Read
   - **Pull Request Threads**: Read & Write
5. Store the token as a secret variable in your Azure DevOps pipeline

### Storing Tokens in Azure DevOps

1. Navigate to **Pipelines** > **Library**
2. Create a new Variable Group or edit an existing one
3. Add the following variables:
   - `GITHUB_PAT` (mark as secret)
   - `AZURE_DEVOPS_PAT` (mark as secret)
4. Link the variable group to your pipeline

Alternatively, you can create the pipeline first and then configure the pipeline-specific variables.

## How It Works

1. **Fetch PR Context**: The task retrieves pull request metadata, existing comments, and iteration details from Azure DevOps
2. **Analyze Changes**: The task identifies all changed files in the most recent iteration
3. **Run Copilot Review**: GitHub Copilot CLI analyzes the changes using the configured or default prompt
4. **Post Comments**: Review findings are posted as comments on the pull request

## Default Review Focus Areas

The default prompt instructs Copilot to focus on:

- **Performance**: Identifying inefficient code patterns
- **Best Practices**: Adherence to coding standards
- **Reusability**: Opportunities for code reuse
- **Maintainability**: Code clarity and documentation
- **Simplification**: Reducing complexity
- **Security**: Potential vulnerabilities
- **Code Consistency**: Style and pattern consistency

## Limitations

- **Windows Only**: Currently requires Windows-based agents
- **GitHub Copilot CLI**: Requires the GitHub Copilot CLI to be installable via `winget`. If using MS-hosted agents, this should be enabled by default.
- **General Comments Only**: Posts general PR comments (file-level inline comments not yet supported)
- **Context Window**: Very large PRs may exceed Copilot's context limits

## Troubleshooting

### Task fails with "GitHub Copilot CLI not found"

Ensure your agent can access `winget` and has internet connectivity to install the Copilot CLI.

### Authentication errors

Verify that:
- Your GitHub PAT has Copilot access
  - If your user account is part of a GitHub organization, ensure the organization admin goes to **GitHub Policies** > **Copilot** > **Copilot CLI** and sets the policy to **Enabled everywhere**
- Your Azure DevOps PAT has Code (Read) and Pull Request Threads (Read & Write) permissions
- Tokens are not expired

### Timeout errors

For large PRs, increase the `timeout` input value. The default is 15 minutes.

### No comments posted

Check the pipeline logs for Copilot's analysis output and determine if the agent experienced connectivity issues when posting comments. Even if Copilot finds no issues, it should still post a single comment indicating as such when using the default prompt.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests on [GitHub](https://github.com/little-fort/ado-copilot-code-review).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For issues and feature requests, please use the [GitHub Issues](https://github.com/little-fort/ado-copilot-code-review/issues) page.

## Acknowledgments

- Built with [Azure Pipelines Task SDK](https://github.com/microsoft/azure-pipelines-task-lib)
- Powered by [GitHub Copilot](https://github.com/features/copilot)
