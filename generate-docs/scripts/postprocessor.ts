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

interface UsedByReference {
    uid: string;          // Full canonical reference UID of the using member
    name: string;         // Display name (e.g., "ExcelScript.Workbook.getWorksheet")
    memberType: string;   // "property", "method", etc.
    packageName: string;  // Package name (e.g., "ExcelScript")
    contextText: string;  // Human-readable context (e.g., "property type")
}

interface UsedByIndex {
    [typeUID: string]: UsedByReference[];
}

interface ApiJsonToken {
    kind: string;
    text?: string;
    canonicalReference?: string;
}

interface ApiJsonMember {
    kind: string;
    canonicalReference?: string;
    name?: string;
    excerptTokens?: ApiJsonToken[];
    members?: ApiJsonMember[];
    returnTypeTokenRange?: { startIndex: number; endIndex: number };
    propertyTypeTokenRange?: { startIndex: number; endIndex: number };
    parameters?: Array<{
        parameterName: string;
        parameterTypeTokenRange: { startIndex: number; endIndex: number };
    }>;
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

// ---- Used By Index Building ----

/**
 * Builds a reverse index mapping type UIDs to the members that use them.
 * Reads the API JSON files produced by api-extractor for each namespace.
 */
function buildUsedByIndex(): UsedByIndex {
    const index: UsedByIndex = {};
    const jsonFiles = [
        path.resolve("../json/json-preview/ExcelScript.api.json"),
        path.resolve("../json/json-preview/OfficeScript.api.json")
    ];

    for (const jsonPath of jsonFiles) {
        if (!fsx.existsSync(jsonPath)) {
            console.warn(`  Warning: JSON file not found: ${jsonPath}`);
            continue;
        }
        try {
            const apiJson: ApiJsonMember = JSON.parse(fsx.readFileSync(jsonPath, 'utf-8'));
            const packageName = path.basename(jsonPath, '.api.json');
            analyzeMembers(apiJson, index, packageName);
        } catch (error) {
            console.warn(`  Warning: Failed to process ${jsonPath}: ${error}`);
        }
    }

    return index;
}

/**
 * Recursively analyzes all members in an API JSON structure.
 */
function analyzeMembers(container: ApiJsonMember, index: UsedByIndex, packageName: string): void {
    if (!container || !container.members) {
        return;
    }

    for (const member of container.members) {
        if (!member.canonicalReference) {
            continue;
        }
        analyzeMemberTypeReferences(member, index, packageName);
        analyzeMembers(member, index, packageName);
    }
}

/**
 * Analyzes a single member for type references in property types,
 * method return types, and method parameter types.
 */
function analyzeMemberTypeReferences(member: ApiJsonMember, index: UsedByIndex, packageName: string): void {
    if (!member.excerptTokens || !member.canonicalReference) {
        return;
    }

    if (member.kind === 'PropertySignature' || member.kind === 'Property') {
        const range = member.propertyTypeTokenRange;
        const tokens = range
            ? member.excerptTokens.slice(range.startIndex, range.endIndex)
            : member.excerptTokens;
        analyzeExcerptTokens(tokens, member.canonicalReference, 'property type', index, packageName);
    }

    if (member.kind === 'MethodSignature' || member.kind === 'Method' || member.kind === 'Function') {
        if (member.returnTypeTokenRange) {
            const returnTokens = member.excerptTokens.slice(
                member.returnTypeTokenRange.startIndex,
                member.returnTypeTokenRange.endIndex
            );
            analyzeExcerptTokens(returnTokens, member.canonicalReference, 'return type', index, packageName);
        }

        if (member.parameters) {
            for (const param of member.parameters) {
                const paramTokens = member.excerptTokens.slice(
                    param.parameterTypeTokenRange.startIndex,
                    param.parameterTypeTokenRange.endIndex
                );
                analyzeExcerptTokens(paramTokens, member.canonicalReference, 'parameter type', index, packageName);
            }
        }
    }
}

/**
 * Scans excerpt tokens for Reference tokens and adds entries to the index.
 */
function analyzeExcerptTokens(
    tokens: ApiJsonToken[],
    usingMemberUID: string,
    contextText: string,
    index: UsedByIndex,
    packageName: string
): void {
    for (const token of tokens) {
        if (token.kind === 'Reference' && token.canonicalReference) {
            // Skip built-in types (e.g., Promise, Array) which start with '!'
            if (token.canonicalReference.startsWith('!')) {
                continue;
            }

            if (!index[token.canonicalReference]) {
                index[token.canonicalReference] = [];
            }

            index[token.canonicalReference].push({
                uid: usingMemberUID,
                name: formatDisplayName(usingMemberUID),
                memberType: getMemberType(usingMemberUID),
                packageName: packageName,
                contextText: contextText
            });
        }
    }
}

/**
 * Converts a canonical reference UID to a human-readable display name.
 * E.g., "ExcelScript!ExcelScript.Workbook#getWorksheet:member(1)" -> "ExcelScript.Workbook.getWorksheet"
 */
function formatDisplayName(uid: string): string {
    const withoutPackage = uid.replace(/^[^!]+!/, '');
    return withoutPackage.replace(/#/g, '.').replace(/:[a-z]+(\(\d+\))?$/, '');
}

/**
 * Determines the member type from a UID string.
 */
function getMemberType(uid: string): string {
    if (uid.includes(':member')) { return 'property'; }
    if (uid.includes(':method')) { return 'method'; }
    if (uid.includes(':function')) { return 'function'; }
    if (uid.includes(':interface')) { return 'interface'; }
    if (uid.includes(':class')) { return 'class'; }
    if (uid.includes(':enum')) { return 'enum'; }
    if (uid.includes(':type')) { return 'type'; }
    return 'member';
}

/**
 * Converts a canonical reference UID to a documentation URL for office-scripts.
 * E.g., "ExcelScript!ExcelScript.Workbook#getWorksheet:member(1)"
 *    -> "/javascript/api/office-scripts/excelscript/excelscript.workbook#excelscript-excelscript-workbook-getworksheet-member(1)"
 */
function convertUidToUrl(uid: string): string {
    const overloadMatch = uid.match(/\((\d+)\)$/);
    const overloadSuffix = overloadMatch ? `(${overloadMatch[1]})` : '';
    const cleanUid = uid.replace(/\(\d+\)$/, '');

    const parts = cleanUid.split('!');
    if (parts.length !== 2) {
        return '';
    }

    const packageName = parts[0].toLowerCase(); // "excelscript" or "officescript"
    const reference = parts[1];

    let classPath: string;
    let memberPart: string;

    const hashIndex = reference.indexOf('#');
    if (hashIndex > 0) {
        classPath = reference.substring(0, hashIndex);
        memberPart = reference.substring(hashIndex + 1);
    } else {
        const colonIndex = reference.indexOf(':');
        if (colonIndex > 0) {
            const beforeColon = reference.substring(0, colonIndex);
            const lastDotIndex = beforeColon.lastIndexOf('.');
            if (lastDotIndex > 0) {
                classPath = reference.substring(0, lastDotIndex);
                memberPart = reference.substring(lastDotIndex + 1);
            } else {
                classPath = reference;
                memberPart = '';
            }
        } else {
            classPath = reference;
            memberPart = '';
        }
    }

    const classPathLower = classPath.toLowerCase();
    let anchor = '';
    if (memberPart) {
        const anchorBase = `${packageName}-${classPath.replace(/\./g, '-')}-${memberPart.replace(/:/g, '-')}`.toLowerCase();
        anchor = `#${anchorBase}${overloadSuffix}`;
    }

    // Namespace-level members (e.g., OfficeScript.downloadFile) have classPath === packageName.
    // Their page is /javascript/api/office-scripts/<pkg>, not /.../<pkg>/<pkg>.
    if (classPathLower === packageName) {
        return `/javascript/api/office-scripts/${packageName}${anchor}`;
    }

    return `/javascript/api/office-scripts/${packageName}/${classPathLower}${anchor}`;
}

/**
 * Deduplicates method overloads, keeping only the first overload entry.
 */
function deduplicateMethodOverloads(references: UsedByReference[]): UsedByReference[] {
    const seen = new Set<string>();
    const result: UsedByReference[] = [];

    for (const ref of references) {
        const baseUID = ref.uid.replace(/\(\d+\)$/, '');
        if (!seen.has(baseUID)) {
            seen.add(baseUID);
            result.push(ref);
        }
    }

    return result;
}

/**
 * Groups references by their containing class name.
 * E.g., "ExcelScript.Workbook.getWorksheet" -> class is "ExcelScript.Workbook"
 */
function groupByContainingClass(references: UsedByReference[]): Record<string, UsedByReference[]> {
    const grouped: Record<string, UsedByReference[]> = {};

    for (const ref of references) {
        const lastDotIndex = ref.name.lastIndexOf('.');
        const className = lastDotIndex > 0 ? ref.name.substring(0, lastDotIndex) : ref.name;

        if (!grouped[className]) {
            grouped[className] = [];
        }
        grouped[className].push(ref);
    }

    return grouped;
}

/**
 * Generates the markdown for the "Used by" section.
 * Groups references by containing class with inline linked member lists.
 */
function generateUsedBySection(references: UsedByReference[]): string {
    if (references.length === 0) {
        return '';
    }

    const lines: string[] = ['\n\n#### Used by\n'];
    const groupedByClass = groupByContainingClass(references);
    const classNames = Object.keys(groupedByClass).sort();

    for (const className of classNames) {
        const members = groupedByClass[className];

        members.sort((a, b) => {
            const memberA = a.name.substring(className.length + 1);
            const memberB = b.name.substring(className.length + 1);
            return memberA.localeCompare(memberB);
        });

        const memberLinks: string[] = [];
        let classUrl = '';

        for (const ref of members) {
            const url = convertUidToUrl(ref.uid);
            const memberName = ref.name.substring(className.length + 1);
            memberLinks.push(`[${memberName}](${url})`);

            if (!classUrl) {
                const uidParts = ref.uid.split('!');
                if (uidParts.length === 2) {
                    const pkgName = uidParts[0].toLowerCase();
                    const classNameLower = className.toLowerCase();
                    // Namespace-level class (e.g., "OfficeScript") has classNameLower === pkgName.
                    // Its page is /javascript/api/office-scripts/<pkg>, not /.../<pkg>/<pkg>.
                    if (classNameLower === pkgName) {
                        classUrl = `/javascript/api/office-scripts/${pkgName}`;
                    } else {
                        classUrl = `/javascript/api/office-scripts/${pkgName}/${classNameLower}`;
                    }
                }
            }
        }

        const classLink = classUrl ? `[${className}](${classUrl})` : className;
        lines.push(`- ${classLink}: ${memberLinks.join(', ')}`);
    }

    return lines.join('\n');
}

/**
 * Injects a "Used by" section into a YAML item's remarks if the item's UID
 * appears in the usedByIndex. Filters out self-references.
 */
function injectUsedBySection(item: any, usedByIndex: UsedByIndex): boolean {
    if (!item.uid) {
        return false;
    }

    const references = usedByIndex[item.uid];
    if (!references || references.length === 0) {
        return false;
    }

    if (!item.remarks) {
        item.remarks = '';
    }

    // Remove any existing "Used by" section before injecting a fresh one
    if (item.remarks.includes('#### Used by')) {
        item.remarks = item.remarks.replace(/\n*#### Used by\n[\s\S]*?(?=\n####|\n*$)/g, '');
    }

    // Extract class name from current item's UID and filter out self-references
    let filteredReferences = references;
    const uidParts = item.uid.split('!');
    if (uidParts.length === 2) {
        const currentClassName = uidParts[1].split(':')[0].split('#')[0];
        filteredReferences = references.filter(ref => {
            const lastDotIndex = ref.name.lastIndexOf('.');
            const refClassName = lastDotIndex > 0 ? ref.name.substring(0, lastDotIndex) : ref.name;
            return refClassName !== currentClassName;
        });
    }

    if (filteredReferences.length === 0) {
        return true;
    }

    const deduplicated = deduplicateMethodOverloads(filteredReferences);
    const usedBySection = generateUsedBySection(deduplicated);

    if (item.remarks.includes('#### Examples')) {
        item.remarks = item.remarks.replace('#### Examples', usedBySection + '\n#### Examples');
    } else if (item.remarks.trim()) {
        item.remarks += usedBySection;
    } else {
        item.remarks = usedBySection.substring(2); // Remove leading \n\n
    }

    return true;
}

tryCatch(async () => {
    console.log("\nStarting postprocessor script...");

    const docsSource = path.resolve("../yaml/yaml-preview");
    const docsDestination = path.resolve("../../docs/docs-ref-autogen");

    // Load snippets from json-preview directory
    const snippetsPath = path.resolve("../json/json-preview/snippets.yaml");
    const { snippetsAll, snippetsTracking } = loadSnippets(snippetsPath);

    // Build the "Used By" reverse index from API JSON files
    console.log("\nBuilding 'Used by' index...");
    const usedByIndex = buildUsedByIndex();
    console.log(`Built 'Used by' index with ${Object.keys(usedByIndex).length} referenced types.`);

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
                fsx.writeFileSync(subFileName, cleanUpYmlFile(ymlFile, snippetsAll, snippetsTracking, usedByIndex));
            });
        } else if (fileName.indexOf("toc") < 0 && fileName.indexOf(".yml") > 0) {
            const ymlFile = fsx.readFileSync(fileName, "utf8");
            fsx.writeFileSync(fileName, cleanUpYmlFile(ymlFile, snippetsAll, snippetsTracking, usedByIndex));
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
    snippetsTracking: SnippetMap,
    usedByIndex: UsedByIndex
): string {
    const schemaComment = ymlFile.substring(0, ymlFile.indexOf("\n") + 1);
    const apiYaml: ApiYaml = jsyaml.load(ymlFile) as ApiYaml;

    // Process the entire YAML tree for examples and API set links
    processYamlItem(apiYaml, snippetsAll, snippetsTracking);

    // Inject "Used by" sections into the top-level item and its members.
    // Enum fields are excluded because OPS doesn't support remarks on enum fields.
    injectUsedBySection(apiYaml, usedByIndex);
    if (apiYaml.properties) {
        apiYaml.properties.forEach(prop => injectUsedBySection(prop, usedByIndex));
    }
    if (apiYaml.methods) {
        apiYaml.methods.forEach(method => injectUsedBySection(method, usedByIndex));
    }

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