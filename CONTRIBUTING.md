# Contributing to Office Studio

Thanks for helping improve Office Studio. Useful contributions include reproducible bug reports, clearer instructions, accessibility improvements, and focused fixes.

## Proposing a change

1. Check existing issues and pull requests for related work.
2. For a substantial feature, open an issue to discuss the intended behavior first.
3. Fork the repository and create a branch for your change.
4. Keep the change focused and explain what it fixes.
5. Run the relevant checks, then submit a pull request against `main`.

The maintainer reviews changes before merging them. Opening an issue or pull request does not require direct write access to this repository. Do not add credentials, personal logs, Office payloads, license keys, or generated executable files to a pull request.

## Building and checking your work

Use Windows with the .NET Framework 4.8 Developer Pack. From the repository root:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Launcher\Build.ps1 -Test
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Launcher\Verify-Portable.ps1
```

UI checks need an interactive Windows desktop. Offline worker tests substitute fake downloads and installers. They must never install or remove Office from the contributor's PC.

For UI changes, include a screenshot and check a smaller window. For startup or asynchronous changes, test the normal application launch as well as self-test mode. Describe any checks you could not run.

## Deployment and removal changes

Preserve the separation between download and installation. Keep cancellation limited to the download's owned processes. Preserve explicit confirmation for Office removal, package validation, Microsoft signature checks, and checks for conflicting installations.

Use Microsoft's supported deployment and removal tools. Do not add activation bypasses, permanent execution-policy changes, unrelated file cleanup, or changes that reduce Office document security settings.

Explain whether verification used mocked packages or actual Microsoft setup. Never describe a mocked test as an end-to-end installation test.

## Reporting problems

Include the Office Studio version, Windows version, selected edition/apps/architecture, reproduction steps, and exact error text. Remove account names, personal paths, license keys, and other sensitive information from logs before posting.

Please keep discussions respectful and focused on the issue.
