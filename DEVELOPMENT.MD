# ExcelFast Development

## Quick Start

1. Clone the repository
2. Open in VS Code with Dev Containers extension installed
3. Let VS Code set up the development container

Alternatively, click "Open in GitHub Codespaces" to start coding immediately.

## Development Environment

ExcelFast uses Development Containers (devcontainers) to ensure consistent development environments across all contributors. The easiest way to get started is using GitHub Codespaces, which automatically sets up the development environment in the cloud.

## Building

Run `.\build.ps1` in PowerShell. This will build the module and output it to the `Build` folder, with an appropriate version.

## Solution Structure

ExcelFast uses a non-traditional C# project structure that [centralizes build artifacts](https://learn.microsoft.com/en-us/dotnet/core/sdk/artifacts-output) and uses longer Pascal-cased folder names for classic paths such as src, bld, etc.:

```
ExcelFast/
├── Artifacts/       # Centralized Build Artifacts, including intermediate .net DLLs, the Published Module, and Nupkg
├── Source/          # Source code
|-- Test/            # Tests for the module.
```

## Architecture

ExcelFast utilizes two separate C# libraries for Excel Manipulation:
`MiniExcel`: High Performance Low memory library that powers Import/Export Workbook
`ClosedXML`: Feature-Rich library for editing Excel files that powers Open/Save Workbook

We must keep the separation distinct as their intents are different.