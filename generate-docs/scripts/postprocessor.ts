#!/usr/bin/env node --harmony

import { generateEnumList } from './util';
import * as fsx from 'fs-extra';
import * as jsyaml from "js-yaml";
import * as path from "path";
import * as os from "os";

const EOL = os.EOL;

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

tryCatch(async () => {
    console.log("\nStarting postprocessor script...");

    const docsSource = path.resolve("../yaml");
    const docsDestination = path.resolve("../../docs/docs-ref-autogen");

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

    // Remove the example field from the YAML as the OPS schema does not support it.
    fsx.readdirSync(docsDestination).forEach(topLevel => { // contents of docs-ref-autogen
        let fileName = docsDestination + '/' + topLevel;
        if (fileName.indexOf(".") < 0) {
            // contents of docs-ref-autogen/<host>
            fsx.readdirSync(fileName)
                .filter(interfaceYml => interfaceYml.indexOf(".yml") >= 0)
                .forEach(interfaceYml => { // contents of docs-ref-autogen/<host>/<host>script
                let subFileName = fileName + '/' + interfaceYml;
                const ymlFile = fsx.readFileSync(subFileName, "utf8");
                fsx.writeFileSync(subFileName, cleanUpYmlFile(ymlFile)); 
            });
        } else if (fileName.indexOf("toc") < 0 && fileName.indexOf(".yml") > 0) {
            const ymlFile = fsx.readFileSync(fileName, "utf8");
            fsx.writeFileSync(fileName, cleanUpYmlFile(ymlFile));
        }
    });

    // Fix all the TOC files.
    console.log("Writing TOC for Office Scripts");
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


function cleanUpYmlFile(ymlFile: string): string {
    const schemaComment = ymlFile.substring(0, ymlFile.indexOf("\n") + 1);
    const apiYaml: ApiYaml = jsyaml.load(ymlFile) as ApiYaml;

    // Add links for type aliases.
    if (apiYaml.uid.endsWith(":type")) {
        let remarks = `${EOL}${EOL}Learn more about the types in this type alias through the following links. ${EOL}${EOL}`
        apiYaml.syntax.substring(apiYaml.syntax.indexOf('=')).match(/[\w]+/g).forEach((match, matchIndex, matches) => {
            remarks += `[ExcelScript.${match}](/javascript/api/office-scripts/excelscript/excelscript.${match.toLowerCase()})`;
            if (matchIndex < matches.length - 1) {
                remarks += ", ";
            }
        });

        let exampleIndex = apiYaml.remarks.indexOf("#### Examples");
        if (exampleIndex > 0) {
            apiYaml.remarks = `${apiYaml.remarks.substring(0, exampleIndex)}${remarks}${EOL}${EOL}${apiYaml.remarks.substring(exampleIndex)}`;
        } else {
            apiYaml.remarks += remarks;
        }
    }
    
    let cleanYml = schemaComment + jsyaml.dump(apiYaml);
    return cleanYml.replace(/^\s*example: \[\]\s*$/gm, "") // Remove example field from yml as the OPS schema does not support it.
                   .replace(/description: \\\*[\r\n]/gm, "description: ''") // Remove descriptions that are just "\*".
                   .replace(/\\\*/gm, "*"); // Fix asterisk protection.
}