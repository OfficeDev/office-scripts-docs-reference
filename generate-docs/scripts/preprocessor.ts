#!/usr/bin/env node --harmony

import * as path from "path";
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";

tryCatch(async () => {
    console.log("\nStarting preprocessor script...");

    const localDtsPath = "../script-inputs/office-scripts-docs.d.ts";

    // Read the combined file and split it into namespace sections.
    console.log(`\nReading combined file from ${path.resolve(localDtsPath)}`);
    let combinedContent = fsx.readFileSync(localDtsPath).toString();

    // Split the content into namespace sections.
    const namespaceSections = splitCombinedContent(combinedContent);

    // Process the ExcelScript section.
    if (!namespaceSections.ExcelScript) {
        throw new Error("ExcelScript namespace section not found");
    }
    let dts = cleanUpDtsContent(namespaceSections.ExcelScript);
    console.log("\ncreate file: excelscript.d.ts (default)");
    fsx.writeFileSync('../api-extractor-inputs-excelscript/excelscript.d.ts', dts);

    // Process the OfficeScript section.
    if (!namespaceSections.OfficeScript) {
        throw new Error("OfficeScript namespace section not found");
    }
    let officeDts = cleanUpDtsContent(namespaceSections.OfficeScript);
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

    readySnippets(`${snippetsSourcePath}/samples.yaml`, "../json/json-preview/snippets.yaml");
    readySnippets(`${snippetsSourcePath}/samples.yaml`, "../json/json-release/snippets.yaml");

    console.log("\nPreprocessor script complete!");
    process.exit(0);
});

function splitCombinedContent(combinedContent: string): Record<string, string> {
    const namespacePattern = /Begin (\w+) namespace/g;
    
    // Find all namespace markers.
    const namespaceMarkers: { name: string, startPos: number }[] = [];
    let match;
    
    while ((match = namespacePattern.exec(combinedContent)) !== null) {
        namespaceMarkers.push({
            name: match[1], // The namespace name (e.g., "ExcelScript", "OfficeScript").
            startPos: match.index
        });
    }
    
    if (namespaceMarkers.length === 0) {
        throw new Error("No namespace markers found in the combined content");
    }
    
    console.log(`Found ${namespaceMarkers.length} namespace(s): ${namespaceMarkers.map(m => m.name).join(', ')}`);
    
    const namespaceSections: Record<string, string> = {};
    
    for (let i = 0; i < namespaceMarkers.length; i++) {
        const currentMarker = namespaceMarkers[i];
        const nextMarker = namespaceMarkers[i + 1];
        
        // Find the end of the current namespace's comment block.
        const commentBlockEnd = combinedContent.indexOf('////////////////////////////////////////////////////////////////', currentMarker.startPos + currentMarker.name.length);
        const contentStart = commentBlockEnd !== -1 ? 
            combinedContent.indexOf('\n', commentBlockEnd) + 1 : 
            combinedContent.indexOf('\n', currentMarker.startPos) + 1;
        
        let contentEnd: number;
        if (nextMarker) {
            // Find the start of the next namespace's comment block.
            contentEnd = combinedContent.lastIndexOf('////////////////////////////////////////////////////////////////', nextMarker.startPos);
        } else {
            // This is the last namespace, so go to the end of the file.
            contentEnd = combinedContent.length;
        }
        
        // Extract the content for this namespace.
        const namespaceContent = combinedContent.substring(contentStart, contentEnd).trim();
        namespaceSections[currentMarker.name] = namespaceContent;
        
        console.log(`Extracted ${currentMarker.name} namespace content (${namespaceContent.length} characters)`);
    }
    
    return namespaceSections;
}

function cleanUpDtsContent(content: string): string {
    console.log("\nFixing issues with d.ts content...");
    return applyRegularExpressions(content);
}

// ----
// Helper function to apply regular expressions to d.ts file contents.
// ----
function applyRegularExpressions (definitionsIn) {
    return definitionsIn.replace(/^(\s*)(declare namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(declare module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(namespace)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(class)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(interface)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(module)(\s+)/gm, `$1export $2$3`)
        .replace(/^(\s*)(function)(\s+)/gm, `$1export $2$3`)
        .replace(/(\s*)(@param)(\s+)(\w+)(\s+)([^\-])/g, `$1$2$3$4$5- $6`)
        .replace(/(\s*)\*(\s*)(@throws)(\s+)(\w+)(.*)/g, `$1$1*$2**Throws**: $4\`$5\`$6`);
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
