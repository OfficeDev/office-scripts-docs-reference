# Office Scripts API Documentation Tools

These tools help support the Office SCripts documentation and the team behind it. Follow these instructions to run the tools in this folder.

## coverage-tester

This tool gives an overview of the documentation coverage for each API. Each API is assessed for documentation quality and the presence of sample code. The quality metrics are still in development.

The output of this tool is a `.csv` file.

### coverage-tester Instructions

1. Clone or fork the repo.
1. In a command window, go to `/office-scripts-docs-reference/generate-docs/tools`
1. Run `npm install`
1. Run `npm run build`
1. Run `node coverage-tester`
1. Open “API Coverage Report.csv”
