#!/usr/bin/env node --harmony

import { generateEnumList } from './util';
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";
import * as path from "path";

interface Toc {
    items: [
        {
            name: string,
            href?: string,
            items: [
                {
                    name: string,
                    uid: string,
                    items: [
                        {
                            name: string,
                            items: [
                                {
                                    name: string,
                                    uid?: string
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
}

interface IMembers {
    items: [
        {
            name: string,
            uid?: string
        }
    ]
}

interface ApiFieldYaml {
    name: string;
    uid: string;
    package: string;
    summary: string;
    remarks?: string;
}

interface ApiPropertyYaml {
    name: string;
    uid: string;
    package: string;
    fullName: string;
    summary: string;
    remarks?: string;
    isPreview: boolean;
    isDeprecated: boolean;
    syntax: {
        content: string;
        return: {
            type: string;
            description?: string;
        }
    }
}

interface ApiMethodYaml {
    name: string;
    uid: string;
    package: string;
    fullName: string;
    summary: string;
    remarks?: string;
    isPreview: boolean;
    isDeprecated: boolean;
    syntax: {
        content: string;
        parameters?: {
            id: string;
            description: string;
            type: string;
        }[];
        return: {
            type: string;
            description: string;
        };
    };
}

interface ApiYaml {
    name: string;
    uid: string;
    package: string;
    fullName: string;
    summary: string;
    remarks: string;
    isPreview: boolean;
    isDeprecated: boolean;
    type: string;
    fields?: ApiFieldYaml[];
    properties?: ApiPropertyYaml[];
    methods?: ApiMethodYaml[];
    syntax?: string;
}

interface SnippetMap {
    [key: string]: string[];
}

/**
 * Loads snippets from the snippets.yaml file
 * @param snippetsPath - Absolute path to snippets.yaml file
 * @returns Object mapping UIDs to snippet arrays and tracking object
 */
function loadSnippets(snippetsPath: string): {
    snippetsAll: SnippetMap,
    snippetsTracking: SnippetMap
} {
    console.log(`Loading snippets from: ${snippetsPath}`);

    if (!fsx.existsSync(snippetsPath)) {
        console.warn(`Warning: Snippets file not found at ${snippetsPath}`);
        return { snippetsAll: {}, snippetsTracking: {} };
    }

    const snippetsContent = fsx.readFileSync(snippetsPath, 'utf8');
    const snippetsAll = jsyaml.load(snippetsContent) as SnippetMap;
    const snippetsTracking = { ...snippetsAll };

    return { snippetsAll, snippetsTracking };
}

/**
 * Extracts the UID without package prefix, matching OfficeYamlDocumenter logic
 * @param uid - Full UID like "ExcelScript!ExcelScript.Application#calculate:member(1)"
 * @returns UID without package like "ExcelScript.Application#calculate:member(1)"
 */
function extractUidWithoutPackage(uid: string): string {
    return uid.replace(/^[^.]+\!/, '');
}

/**
 * Generates the example text from snippets array
 * Matches OfficeYamlDocumenter._generateExampleSnippetText (lines 130-136)
 * @param snippets - Array of code snippet strings
 * @returns Formatted example text
 */
function generateExampleSnippetText(snippets: string[]): string {
    const text = ['\n\n#### Examples\n'];
    for (const snippet of snippets) {
        text.push(`\`\`\`TypeScript\n${snippet}\n\`\`\``);
    }
    return text.join('\n');
}

/**
 * Inserts examples into a YAML item
 * Matches OfficeYamlDocumenter.onCustomizeYamlItem logic (lines 93-109)
 * @param yamlItem - The YAML item to modify
 * @param snippetText - The formatted example text to insert
 */
function insertExamples(yamlItem: any, snippetText: string): void {
    // Priority 1: Append to remarks if it exists (lines 97-98)
    if (yamlItem.remarks) {
        yamlItem.remarks += snippetText;
    }
    // Priority 2: Append to syntax.return.description if no remarks (lines 100-104)
    else if (yamlItem.syntax && yamlItem.syntax.return) {
        if (!yamlItem.syntax.return.description) {
            yamlItem.syntax.return.description = '';
        }
        yamlItem.syntax.return.description += snippetText;
    }
    // Priority 3: Create new remarks (lines 106-107)
    else {
        yamlItem.remarks = snippetText;
    }
}

/**
 * Fixes API set references in markup text
 * Matches OfficeYamlDocumenter._fixupApiSet (lines 111-119)
 * @param markup - The text to process
 * @param uid - The UID for determining product context
 * @returns Fixed markup
 */
function fixupApiSet(markup: string | undefined, uid: string): string | undefined {
    if (!markup) {
        return markup;
    }

    // Change "Api" to "API" (line 117)
    markup = markup.replace(/Api/g, 'API');

    // Wrap in markdown link (line 118)
    // Pattern: \[API set: ExcelApi 1.1\] -> \[ [API set: ExcelApi 1.1](url) \]
    const apiSetUrl = getApiSetUrl(uid);
    return markup.replace(/\\\[(API set:[^\]]+)\\\]/g, `\\[ [$1](${apiSetUrl}) \\]`);
}

/**
 * Gets the API set documentation URL based on product context
 * Matches OfficeYamlDocumenter._getApiSetUrl (lines 120-129)
 * @param uid - The UID to check for product context
 * @returns The appropriate API set documentation URL
 */
function getApiSetUrl(uid: string): string {
    const apiSetUrls: { [key: string]: string } = {
        Excel: '/javascript/api/requirement-sets/excel/excel-api-requirement-sets',
        OneNote: '/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets',
        Outlook: '/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets',
        PowerPoint: '/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets',
        Visio: '/office/dev/add-ins/reference/overview/visio-javascript-reference-overview',
        Word: '/javascript/api/requirement-sets/word/word-api-requirement-sets'
    };

    // Check for product match (case-insensitive)
    for (const key of Object.keys(apiSetUrls)) {
        const regexp = new RegExp(key, 'i');
        if (regexp.test(uid)) {
            return apiSetUrls[key];
        }
    }

    // Default URL (line 128)
    return '/office/dev/add-ins/reference/javascript-api-for-office';
}

/**
 * Recursively processes YAML items to insert examples and fix API set links
 * @param item - The YAML item or array to process
 * @param snippetsAll - All snippets for lookup
 * @param snippetsTracking - Tracking object to mark used snippets
 */
function processYamlItem(
    item: any,
    snippetsAll: SnippetMap,
    snippetsTracking: SnippetMap
): void {
    if (!item || typeof item !== 'object') {
        return;
    }

    // Process arrays
    if (Array.isArray(item)) {
        for (const element of item) {
            processYamlItem(element, snippetsAll, snippetsTracking);
        }
        return;
    }

    // Process object with UID
    if (item.uid && typeof item.uid === 'string') {
        const nameWithoutPackage = extractUidWithoutPackage(item.uid);

        // Fix API set links in summary and remarks (lines 87-91)
        if (item.summary) {
            item.summary = fixupApiSet(item.summary, item.uid);
        }
        if (item.remarks) {
            item.remarks = fixupApiSet(item.remarks, item.uid);
        }

        // Insert examples if snippets exist (lines 93-109)
        const snippets = snippetsAll[nameWithoutPackage];
        if (snippets && snippets.length > 0) {
            // Mark as used (line 95)
            delete snippetsTracking[nameWithoutPackage];

            const snippetText = generateExampleSnippetText(snippets);
            insertExamples(item, snippetText);
        }
    }

    // Recursively process all properties
    for (const key of Object.keys(item)) {
        if (key !== 'uid' && key !== 'summary' && key !== 'remarks' && key !== 'syntax') {
            processYamlItem(item[key], snippetsAll, snippetsTracking);
        }
    }
}

tryCatch(async () => {
    console.log("\nStarting postprocessor script...");

    const docsSource = path.resolve("../yaml/yaml-preview");
    const docsDestination = path.resolve("../../docs/docs-ref-autogen");

    // Load snippets from json-preview directory
    const snippetsPath = path.resolve("../json/json-preview/snippets.yaml");
    const { snippetsAll, snippetsTracking } = loadSnippets(snippetsPath);

    console.log(`Deleting old docs at: ${docsDestination}`);
    // Delete everything except the 'overview' files.
    fsx.readdirSync(docsDestination).forEach(topLevel => {
        if (topLevel.indexOf(".yml") >= 0) {
            fsx.removeSync(docsDestination + '/' + topLevel);
        } else if (topLevel.indexOf(".") < 0) {
            fsx.readdirSync(docsDestination + '/' + topLevel).forEach(filename => {  
                fsx.removeSync(docsDestination + '/' + topLevel + '/' + filename);
            });
        }
    });

    console.log(`Copying docs output files to: ${docsDestination}`);
    // Copy docs output to the /docs/docs-ref-autogen folder.
    fsx.readdirSync(docsSource)
        .forEach(filename => {
            fsx.copySync(
                docsSource + '/' + filename,
                docsDestination + '/' + filename
            );
    });

    // Process YAML files: remove example fields and insert snippets
    console.log("\nProcessing YAML files...");
    fsx.readdirSync(docsDestination).forEach(topLevel => { // Contents of docs-ref-autogen.
        let fileName = docsDestination + '/' + topLevel;
        if (fileName.indexOf(".") < 0) {
            // Contents of docs-ref-autogen/<host>.
            fsx.readdirSync(fileName)
                .filter(interfaceYml => interfaceYml.indexOf(".yml") >= 0)
                .forEach(interfaceYml => { // Contents of docs-ref-autogen/<host>/<host>script.
                let subFileName = fileName + '/' + interfaceYml;
                const ymlFile = fsx.readFileSync(subFileName, "utf8");
                fsx.writeFileSync(subFileName, cleanUpYmlFile(ymlFile, snippetsAll, snippetsTracking));
            });
        } else if (fileName.indexOf("toc") < 0 && fileName.indexOf(".yml") > 0) {
            const ymlFile = fsx.readFileSync(fileName, "utf8");
            fsx.writeFileSync(fileName, cleanUpYmlFile(ymlFile, snippetsAll, snippetsTracking));
        }
    });

    // Log unused snippets (matching OfficeYamlDocumenter behavior)
    console.log("\nChecking for unused snippets...");
    const unusedSnippets = Object.keys(snippetsTracking);
    if (unusedSnippets.length > 0) {
        console.warn(`\nWarning: Found ${unusedSnippets.length} unused snippet(s):`);
        for (const apiName of unusedSnippets) {
            console.warn(`  - ${apiName}`);
        }
    } else {
        console.log("All snippets were used successfully.");
    }

    // Fix all the TOC files.
    console.log("\nWriting TOC for Office Scripts");
    let tocPath = path.resolve(docsDestination) + "/toc.yml";
    let latestToc = fixToc(tocPath);
    fsx.writeFileSync(tocPath, jsyaml.dump(latestToc));


    console.log("\nPostprocessor script complete!\n");
    process.exit(0);
});

async function tryCatch(call: () => Promise<void>) {
    try {
        await call();
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
}

function fixToc(tocPath: string): Toc {
    console.log(`Updating the structure of the TOC file: ${tocPath}`);

    let origToc = (jsyaml.load(fsx.readFileSync(tocPath).toString()) as Toc);
    let newToc = <Toc>{};
    let membersToMove = <IMembers>{};

    newToc.items = [{
        "name": "API reference",
        "items": [] as any
    }];
    newToc.items[0].items = [{
        "name": "API reference overview",
        "href": "overview.md"
    },{
        "name": "Preview APIs",
        "href": "preview-apis.md"
    }] as any;

    // Create a folder for enums.
    origToc.items.forEach((rootItem) => {
        rootItem.items.forEach((packageItem) => {
            let enumFilter = generateEnumList(fsx.readFileSync(`../api-extractor-inputs-${packageItem.name}/${packageItem.name}.d.ts`).toString());
            membersToMove.items = packageItem.items;

            let enumList = membersToMove.items.filter(item => {
                return enumFilter.indexOf(item.name) >= 0;
            });
            let primaryList = membersToMove.items.filter(item => {
                return enumFilter.indexOf(item.name) < 0;
            });

            let enumRoot = {"name": "Enums", "uid": "", "items": enumList};
            primaryList.unshift(enumRoot);
            newToc.items[0].items.push({
                "name": packageItem.name,
                "uid": packageItem.uid,
                "items": primaryList as any
            });
        });
    });

    return newToc;
}


function cleanUpYmlFile(
    ymlFile: string,
    snippetsAll: SnippetMap,
    snippetsTracking: SnippetMap
): string {
    const schemaComment = ymlFile.substring(0, ymlFile.indexOf("\n") + 1);
    const apiYaml: ApiYaml = jsyaml.load(ymlFile) as ApiYaml;

    // Process the entire YAML tree for examples and API set links
    processYamlItem(apiYaml, snippetsAll, snippetsTracking);

    // Add links for type aliases.
    if (apiYaml.uid.endsWith(":type")) {
        let remarks = `\n\nLearn more about the types in this type alias through the following links. \n\n`
        apiYaml.syntax.substring(apiYaml.syntax.indexOf('=')).match(/[\w]+/g).forEach((match, matchIndex, matches) => {
            remarks += `[ExcelScript.${match}](/javascript/api/office-scripts/excelscript/excelscript.${match.toLowerCase()})`;
            if (matchIndex < matches.length - 1) {
                remarks += ", ";
            }
        });

        let exampleIndex = apiYaml.remarks.indexOf("#### Examples");
        if (exampleIndex > 0) {
            apiYaml.remarks = `${apiYaml.remarks.substring(0, exampleIndex)}${remarks}\n\n${apiYaml.remarks.substring(exampleIndex)}`;
        } else {
            apiYaml.remarks += remarks;
        }
    }
    
    let cleanYml = schemaComment + jsyaml.dump(apiYaml);
    return cleanYml.replace(/^\s*example: \[\]\s*$/gm, "") // Remove example field from yml as the OPS schema does not support it.
                   .replace(/description: \\\*[\r\n]/gm, "description: ''") // Remove descriptions that are just "\*".
                   .replace(/\\\*/gm, "*"); // Fix asterisk protection.
}