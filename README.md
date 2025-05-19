# xcelToBDDSeliniumCodeGenerator
# Excel to SpecFlow Selenium Code Generator

A console application that reads test case steps from an Excel file and generates Selenium WebDriver automation code structured by Elements Class, Page Class, and Step Definition Class for SpecFlow in C#.

## Features

- Parses Excel sheets with test steps.
- Generates reusable Elements, Page, and Step Definition classes.
- Supports SpecFlow BDD style automation framework structure.
- Helps speed up automation test development from manual test cases.

## Getting Started

### Prerequisites

- [.NET 6.0 SDK or later](https://dotnet.microsoft.com/download)
- Excel file with test steps formatted as:

| Step Description      | Element Locator | Action | Input Data |
|-----------------------|-----------------|--------|------------|
| Click Login Button     | id=loginBtn     | Click  |            |
| Enter Username        | id=username     | SendKeys | testuser  |
| ...                   | ...             | ...    | ...        |

### How to Use

1. Clone the repo:
   ```bash
   git clone https://github.com/yourusername/ExcelToSpecFlowCodeGenerator.git
   cd ExcelToSpecFlowCodeGenerator
