#!/usr/bin/env node --harmony

import * as path from "path";
import * as fsx from 'fs-extra';
import yaml = require('js-yaml');

tryCatch(async () => {
    console.log("\nStarting preprocessor script...");

    const localDtsPath = "../script-inputs/excel.d.ts";
    const localAsyncDtsPath = "../script-inputs/excel-async.d.ts";

    let dts = cleanUpDts(localDtsPath);
    let asyncDts = handleLiteralParameterOverloads(cleanUpDts(localAsyncDtsPath));

    console.log("\ncreate file: excel.d.ts (default)");
    fsx.writeFileSync('../api-extractor-inputs-excel/excel.d.ts', dts);

    console.log("\ncreate file: excel-async.d.ts");
    fsx.writeFileSync('../api-extractor-inputs-excel-async/excel.d.ts', asyncDts);

    // ----
    // Process Snippets
    // ----
    console.log("\nRemoving old snippets input files...");
    const scriptInputsPath = path.resolve("../script-inputs");
    fsx.readdirSync(scriptInputsPath)
        .filter(filename => filename.indexOf("snippets") > 0)
        .forEach(filename => fsx.removeSync(scriptInputsPath + '/' + filename));

    console.log("\nCreating snippets file...");
    const snippetsSourcePath = path.resolve("../../docs/code-snippets");
    console.log("\nReading from files: " + snippetsSourcePath);

    let localSnippetsString = fsx.readFileSync(`${snippetsSourcePath}/excel-snippets.yaml`).toString();
    let localAsyncSnippetsString = fsx.readFileSync(`${snippetsSourcePath}/excel-async-snippets.yaml`).toString();

    // Parse the YAML into an object/hash set.
    let snippets = yaml.load(localSnippetsString);
    let asyncSnippets = yaml.load(localAsyncSnippetsString);

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

    let asyncSnippetDestination = path.resolve("../json/excel-async/snippets.yaml");
    console.log("\nWriting snippets to: " + asyncSnippetDestination);
    fsx.createFileSync(asyncSnippetDestination);
    fsx.writeFileSync(asyncSnippetDestination, yaml.safeDump(
        asyncSnippets,
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

function handleLiteralParameterOverloads(dtsString: string): string {
    // rename parameters for string literal overloads
    const matches = dtsString.match(/([a-zA-Z]+)\??: (\"[a-zA-Z]*\").*:/g);
    let matchIndex = 0;
    if (matches) {
        matches.forEach((match) => {
            let parameterName = match.substring(0, match.indexOf(": "));
            matchIndex = dtsString.indexOf(match, matchIndex);
            parameterName = parameterName.indexOf("?") >= 0 ? parameterName.substring(0, parameterName.length - 1) : parameterName;
            const parameterString = "@param " + parameterName + " ";
            const index = dtsString.lastIndexOf(parameterString, matchIndex);
            if (index < 0) {
                console.warn("Missing @param for literal parameter: " + match);
            } else {
            dtsString = dtsString.substring(0, index)
            + "@param " + parameterName + "String "
            + dtsString.substring(index + parameterString.length);
            matchIndex += match.length;
            }
        });
    }
    return dtsString.replace(/([a-zA-Z]+)(\??: \"[a-zA-Z]*\".*:)/g, "$1String$2");
}

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}
