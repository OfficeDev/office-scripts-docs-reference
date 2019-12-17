#!/usr/bin/env node --harmony

// import { promptFromList } from './simple-prompts';
import * as path from "path";
import * as fsx from 'fs-extra';
import yaml = require('js-yaml');
import { fetchAndThrowOnError } from './util';

tryCatch(async () => {
    // ----
    // Display prompt
    // ----
    console.log('\n\n');
    // TODO: Add CDN edipoint link
    // const urlToCopyOfficeJsFrom = await promptFromList({
    //     message: `What is the source of the Office-js TypeScript definition file that should be used to generate the Excel Script docs?`,
    //     choices: [
    //         { name: "Local file [generate-docs\\script-inputs\\excel.d.ts]", value: "" }
    //     ]
    // });

    console.log("\nStarting preprocessor script...");

    const localReleaseDtsPath = "../script-inputs/excel.d.ts";
    // if (urlToCopyOfficeJsFrom.length > 0) {
    //     fsx.writeFileSync(localReleaseDtsPath, await fetchAndThrowOnError(urlToCopyOfficeJsFrom, "text"));
    // }

    let releaseDefinitions = cleanUpDts(localReleaseDtsPath);

    console.log("\ncreate file: excel.d.ts (preview)");
    fsx.writeFileSync(
        '../api-extractor-inputs-excel/excel.d.ts',
        handleLiteralParameterOverloads(excelSpecificCleanup(releaseDefinitions))
    );

    // TODO: Deal with Script Lab snippets
    // ----
    // Process Snippets
    // ----
    console.log("\nRemoving old snippets input files...");

    const scriptInputsPath = path.resolve("../script-inputs");
    fsx.readdirSync(scriptInputsPath)
        .filter(filename => filename.indexOf("snippets") > 0)
        .forEach(filename => fsx.removeSync(scriptInputsPath + '/' + filename));

    console.log("\nCreating snippets file...");
    console.log("\nReading from files: " + path.resolve("../../docs/code-snippets"));

    const snippetsSourcePath = path.resolve("../../docs/code-snippets");
    let localCodeSnippetsString : string = "";
    fsx.readdirSync(path.resolve(snippetsSourcePath))
        .filter(name => name.endsWith('.yaml') || name.endsWith('.yml'))
        .forEach((filename, index) => {
            localCodeSnippetsString += fsx.readFileSync(`${snippetsSourcePath}/${filename}`).toString() + "\r\n";
        });

    fsx.writeFileSync("../script-inputs/local-repo-snippets.yaml", localCodeSnippetsString);

    // Parse the YAML into an object/hash set.
    let snippets = yaml.load(localCodeSnippetsString);

    let snippetDestination = path.resolve("../json/excel/snippets.yaml")
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

function excelSpecificCleanup(dtsContent: string) {
    return dtsContent.replace(/export interface .*Set {\r?\n.*Icon;/gm, `/** [Api set: ExcelApi 1.2] */\n\t$&`)
        .replace("export interface IconCollections {", "/** [Api set: ExcelApi 1.2] */\n\texport interface IconCollections {")
        .replace("var icons: IconCollections;", "/** [Api set: ExcelApi 1.2] */\n\tvar icons: IconCollections;");
}

function cleanUpDts(localDtsPath: string): string {
    console.log(`\nReading from ${path.resolve(localDtsPath)}`);
    let definitions = fsx.readFileSync(localDtsPath).toString();

    console.log("\nFixing issues with d.ts file...");
    return applyRegularExpressions(
        definitions
        .replace(/([ ]*)load\(option\?: string \| string\[\]\): (Excel|Word|OneNote|Visio)\.(.*);/g,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.\n$1 */\n$1load(propertyNames?: string | string[]): $2.$3;")
        .replace(/([ ]*)load\(option\?: {\n[ ]*select\?: string;\n[ ]*expand\?: string;\n[ ]*}\): (Excel|Word|OneNote|Visio)\.(.*);/gm,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.\n$1 */\n$1load(propertyNamesAndPaths?: { select?: string; expand?: string; }): $2.$3;")
        .replace(/([ ]*)load\(option\?: (Excel|Word|OneNote|Visio)\.Interfaces\.(.*)CollectionLoadOptions & [Excel|Word|OneNote|Visio]\.Interfaces\.CollectionLoadOptions\): [Excel|Word|OneNote|Visio]\.[.*]Collection;/g,
                 "$1/**\n$1 * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.\n$1 * @param collectionLoadOptions - Where collectionLoadOptions.select is a comma-delimited string that specifies the properties to load, and collectionLoadOptions.expand is a comma-delimited string that specifies the navigation properties to load. collectionLoadOptions.top specifies the maximum number of collection items that can be included in the result. collectionLoadOptions.skip specifies the number of items that are to be skipped and not included in the result. If collectionLoadOptions.top is specified, the result set will start after skipping the specified number of items.\n$1 */\n$1load(collectionLoadOptions?: $2.Interfaces.$3CollectionLoadOptions & $2.Interfaces.CollectionLoadOptions): $2.$3Collection;")
        .replace(/(extends OfficeCore.RequestContext)/g, `extends OfficeExtension.ClientRequestContext`));
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
