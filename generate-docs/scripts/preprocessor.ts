#!/usr/bin/env node --harmony

import * as path from "path";
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";

tryCatch(async () => {
    console.log("\nStarting preprocessor script...");

    const localDtsPath = "../script-inputs/office-scripts-docs.d.ts";

    // Read the combined file and split it into two sections
    console.log(`\nReading combined file from ${path.resolve(localDtsPath)}`);
    let combinedContent = fsx.readFileSync(localDtsPath).toString();

    // Split the content into ExcelScript and OfficeScript sections
    const { excelScriptContent, officeScriptContent } = splitCombinedContent(combinedContent);

    // Process ExcelScript section
    let dts = cleanUpDtsContent(excelScriptContent);
    console.log("\ncreate file: excelscript.d.ts (default)");
    fsx.writeFileSync('../api-extractor-inputs-excelscript/excelscript.d.ts', dts);

    // Process OfficeScript section
    let officeDts = cleanUpDtsContent(officeScriptContent);
    console.log("\ncreate file: officescript.d.ts");
    fsx.writeFileSync('../api-extractor-inputs-officescript/officescript.d.ts', officeDts);

    // ----
    // Process Snippets
    // ----
    console.log("\nRemoving old snippets input files...");
    const scriptInputsPath = path.resolve("../script-inputs");
    fsx.readdirSync(scriptInputsPath)
        .filter(filename => filename.indexOf("snippets") > 0)
        .forEach(filename => fsx.removeSync(scriptInputsPath + '/' + filename));

    console.log("\nCreating snippets file...");
    const snippetsSourcePath = path.resolve("../../docs/sample-scripts");
    console.log("\nReading from files: " + snippetsSourcePath);

    readySnippets(`${snippetsSourcePath}/samples.yaml`, "../json/snippets.yaml");

    console.log("\nPreprocessor script complete!");
    process.exit(0);
});

function splitCombinedContent(combinedContent: string): { excelScriptContent: string, officeScriptContent: string } {
    const beginExcelScript = "Begin ExcelScript namespace";
    const beginOfficeScript = "Begin OfficeScript namespace";
    
    const excelScriptStart = combinedContent.indexOf(beginExcelScript);
    const officeScriptStart = combinedContent.indexOf(beginOfficeScript);
    
    if (excelScriptStart === -1) {
        throw new Error("Could not find 'Begin ExcelScript namespace' comment block");
    }
    
    if (officeScriptStart === -1) {
        throw new Error("Could not find 'Begin OfficeScript namespace' comment block");
    }
    
    // Find the end of the ExcelScript comment block and start of actual content
    const excelScriptCommentEnd = combinedContent.indexOf('////////////////////////////////////////////////////////////////', excelScriptStart + beginExcelScript.length);
    const excelScriptContentStart = excelScriptCommentEnd !== -1 ? 
        combinedContent.indexOf('\n', excelScriptCommentEnd) + 1 : 
        combinedContent.indexOf('\n', excelScriptStart + beginExcelScript.length) + 1;
    
    // Find the start of the OfficeScript comment block
    const officeScriptCommentStart = combinedContent.lastIndexOf('////////////////////////////////////////////////////////////////', officeScriptStart);
    
    // Extract ExcelScript content (from after ExcelScript comment block until before OfficeScript comment block)
    const excelScriptContent = combinedContent.substring(excelScriptContentStart, officeScriptCommentStart).trim();
    
    // Find the end of the OfficeScript comment block and start of actual content
    const officeScriptCommentEnd = combinedContent.indexOf('////////////////////////////////////////////////////////////////', officeScriptStart + beginOfficeScript.length);
    const officeScriptContentStart = officeScriptCommentEnd !== -1 ? 
        combinedContent.indexOf('\n', officeScriptCommentEnd) + 1 : 
        combinedContent.indexOf('\n', officeScriptStart + beginOfficeScript.length) + 1;
    
    // Extract OfficeScript content (from after OfficeScript comment block to end of file)
    const officeScriptContent = combinedContent.substring(officeScriptContentStart).trim();
    
    return { excelScriptContent, officeScriptContent };
}

function cleanUpDtsContent(content: string): string {
    console.log("\nFixing issues with d.ts content...");
    return applyRegularExpressions(content);
}

// ----
// Helper function to apply regular expressions to d.ts file contents
// ----
function applyRegularExpressions (definitionsIn) {
    return definitionsIn.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s)(\s)/g, `$1$2$3$4$5`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`);
}

function readySnippets(snippetsSourceFile: string, snippetDestinationFile: string) {
    let localSnippetsString = fsx.readFileSync(snippetsSourceFile).toString();

    // Parse the YAML into an object/hash set.
    let snippets = jsyaml.load(localSnippetsString);
    let snippetDestination = path.resolve(snippetDestinationFile);
    console.log("\nWriting snippets to: " + snippetDestination);
    fsx.createFileSync(snippetDestination);
    fsx.writeFileSync(snippetDestination, jsyaml.dump(
        snippets,
        {sortKeys: <any>((a: string, b: string) => {
            if (a < b) {
                return -1;
            } else if (a > b) {
                return 1;
            } else {
                return 0;
            }
        })}
    ));
}

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
