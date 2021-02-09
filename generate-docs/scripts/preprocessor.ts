#!/usr/bin/env node --harmony

import * as path from "path";
import * as fsx from 'fs-extra';
import yaml = require('js-yaml');

tryCatch(async () => {
    console.log("\nStarting preprocessor script...");

    const localDtsPath = "../script-inputs/excel.d.ts";

    let dts = cleanUpDts(localDtsPath);

    console.log("\ncreate file: excel.d.ts (default)");
    fsx.writeFileSync('../api-extractor-inputs-excel/excel.d.ts', dts);

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

    let localSnippetsString = fsx.readFileSync(`${snippetsSourcePath}/excel-scripts.yaml`).toString();

    // Parse the YAML into an object/hash set.
    let snippets = yaml.load(localSnippetsString);
    let snippetDestination = path.resolve("../json/excel/snippets.yaml");
    console.log("\nWriting snippets to: " + snippetDestination);
    fsx.createFileSync(snippetDestination);
    fsx.writeFileSync(snippetDestination, yaml.safeDump(
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

    console.log("\nPreprocessor script complete!");
    process.exit(0);
});

function cleanUpDts(localDtsPath: string): string {
    console.log(`\nReading from ${path.resolve(localDtsPath)}`);
    let definitions = fsx.readFileSync(localDtsPath).toString();

    console.log("\nFixing issues with d.ts file...");
    return applyRegularExpressions(definitions.replace(/(extends OfficeCore.RequestContext)/g, `extends OfficeExtension.ClientRequestContext`));
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

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
