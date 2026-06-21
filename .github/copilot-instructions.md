# Project Coding Standards

## MiniExcel Guidelines
- Use the 2.0 API when possible, but the 1.0 API is still supported and may be used if the 2.0 API does not meet the needs of a particular scenario. Use a comment that begins with HACK: to note any use of the 1.0 API, and include a link to the relevant issue in the GitHub repository if applicable.
- Use the README for implementation guidelines https://raw.githubusercontent.com/mini-software/MiniExcel/refs/heads/master/README-V2.md
- Consider the upgrade notes https://raw.githubusercontent.com/mini-software/MiniExcel/refs/heads/master/V2-Upgrade-Notes.md

## CSharp C# Guidelines

- Use File Scoped Namespaces
- Use Collection Expressions
- Prefer explicit type over var
- Use Linq methods for collections when appropriate
- Use primary constructors when appropriate
- Prefer records for DTOs

## PowerShell Guidelines

- Use splatting if a command with parameters would be longer than 180 characters

## Pester Testing Guidelines

- Write Pester tests for PowerShell cmdlets under Test/PowerShell using Describe, Context, and It blocks.
- Prefer real behavior over mocks; use the shared fixtures in Test/Fixtures/Fixtures.ps1 and real workbook files when possible. Create new workbook file fixtures if needed.
- Cover the main user-facing scenarios: success, invalid input, and expected failures.
- When asserting failures, use `Should -Throw` with `-ErrorAction Stop` and verify the expected `-ExceptionType` or `-ErrorId`.
- Keep tests readable and focused on observable outcomes such as return type, object properties, counts, and error behavior.
- Name tests clearly to describe the scenario and expected result.
- Run tests frequently during development to catch issues early and ensure coverage of new code.
- To run tests, first try `Invoke-Build Test` to run all tests, or `Invoke-Build Test -TestName <test name>` to run a specific test.