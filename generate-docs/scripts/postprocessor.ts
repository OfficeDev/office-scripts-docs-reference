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

tryCatch(async () => {
    console.log("\nStarting postprocessor script...");

    const docsSource = path.resolve("../yaml");
    const docsDestination = path.resolve("../../docs/docs-ref-autogen");

    console.log(`Deleting old docs at: ${docsDestination}`);
    // Delete everything except the 'overview' files.
    fsx.readdirSync(docsDestination).forEach(folder => {
        fsx.readdirSync(docsDestination + '/' + folder).forEach(filename => {
            if (!filename.includes("overview.md")) {
                fsx.removeSync(docsDestination + '/' + folder + '/' + filename);
            }});
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
    fsx.readdirSync(docsDestination)
        .filter(topLevel => topLevel.indexOf(".") < 0)
        .forEach(topLevel => { // contents of docs-ref-autogen
            let hostFolder = docsDestination + '/' + topLevel;
            fsx.readdirSync(hostFolder).forEach((subfilename) => {
                if (subfilename.indexOf(".") < 0) { 
                    // contents of docs-ref-autogen/<host>
                    let scriptFolder = hostFolder + '/' + subfilename;
                    fsx.readdirSync(scriptFolder)
                        .filter(interfaceYml => interfaceYml.indexOf(".yml") >= 0)
                        .forEach(interfaceYml => { // contents of docs-ref-autogen/<host>/<host>script
                        fsx.writeFileSync(
                            scriptFolder + '/' + interfaceYml,
                            fsx.readFileSync(scriptFolder + '/' + interfaceYml).toString().replace(/^\s*example: \[\]\s*$/gm, "").replace(/\\\*/gm, "*")
                        );
                    });
                } else if (subfilename.indexOf("toc") < 0 && subfilename.indexOf(".yml") > 0) {
                    fsx.writeFileSync(
                        hostFolder + '/' + subfilename,
                        fsx.readFileSync(hostFolder + '/' + subfilename).toString().replace(/^\s*example: \[\]\s*$/gm, "").replace(/\\\*/gm, "*")
                    );
                }
        });
    });

    // Fix all the TOC file.
    console.log("Writing TOC for Office Scripts");
    let versionPath = path.resolve(`${docsDestination}/excel`);
    let tocPath = versionPath + "/toc.yml";
    let latestToc = fixToc(tocPath, "../api-extractor-inputs-excel/excel.d.ts");
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

function fixToc(tocPath: string, sourceDtsPath: string): Toc {
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
    let excelEnumFilter = generateEnumList(fsx.readFileSync(sourceDtsPath).toString());
    origToc.items.forEach((rootItem) => {
        rootItem.items.forEach((packageItem) => {
            membersToMove.items = packageItem.items;

            let enumList = membersToMove.items.filter(item => {
                return excelEnumFilter.indexOf(item.name) >= 0;
            });
            let primaryList = membersToMove.items.filter(item => {
                return excelEnumFilter.indexOf(item.name) < 0;
            });

            let excelEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
            primaryList.unshift(excelEnumRoot);
            newToc.items[0].items.push({
                "name": packageItem.name,
                "uid": packageItem.uid,
                "items": primaryList as any
            });
        });
    });

    return newToc;
}
