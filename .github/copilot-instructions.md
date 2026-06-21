# Project Coding Standards

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