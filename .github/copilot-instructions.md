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
- After making changes to the code, run tests in the following order:
  1. Run tests specific to your changes or the cmdlet you changed first to get quick feedback. You can do this by running `Invoke-Build Pester -TestName <TestName>` with a filter for the specific test or tests you want to run. Construct the test name filter using the Describe, Context, and It block names combined with dots. For example, if you have a test defined as `Describe 'New-ExcelFile' { Context 'with valid parameters' { It 'should create a new Excel file' { ... } } }`, you could run just that test with `Invoke-Build Pester -TestName 'New-ExcelFile.with valid parameters.should create a new Excel file'`. This allows you to quickly verify that your changes work as expected before running the full test suite.
  2. After any individual tests you have changed pass, run `Invoke-Build Pester` to verify that the PowerShell cmdlets work as expected and that the Pester tests pass. If a test fails, investigate the failure and try to fix the code first, then fix the test if needed, then re-run the tests to verify that they pass.
  3. If `Invoke-Build Pester` succeeds, run `Invoke-Build Pester-WinPS` if on Windows second to verify that the cmdlets work in Windows PowerShell