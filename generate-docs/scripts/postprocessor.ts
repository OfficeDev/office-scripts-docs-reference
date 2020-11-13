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
    // copy docs output to /docs/docs-ref-autogen folder
    fsx.readdirSync(docsSource)
        .forEach(filename => {
            fsx.copySync(
                docsSource + '/' + filename,
                docsDestination + '/' + filename
            );
    });

    // change the async files to match the sync namespace pattern
    const asyncDocFolder = docsDestination + "/excel-async/excel";
    fsx.readdirSync(asyncDocFolder).forEach(filename => {
        fsx.renameSync(
            asyncDocFolder + '/' + filename,
            asyncDocFolder + '/' + filename.replace("excel.", "excelscript.")
        );
        fsx.removeSync(asyncDocFolder + '/' + filename);
    });
    fsx.renameSync(asyncDocFolder, docsDestination + "/excel-async/excelscript");
    fsx.renameSync(asyncDocFolder + ".yml", docsDestination + "/excel-async/excelscript.yml");

    // fix all the individual TOC files
    console.log("Writing TOC for Office Scripts");
    let versionPath = path.resolve(`${docsDestination}/excel`);
    let tocPath = versionPath + "/toc.yml";
    let latestToc = fixToc(tocPath, "../api-extractor-inputs-excel/excel.d.ts");
    fsx.writeFileSync(tocPath, jsyaml.safeDump(latestToc));

    console.log("Writing TOC for Office Scripts Async");
    versionPath = path.resolve(`${docsDestination}/excel-async`);
    const asyncTocPath = versionPath + "/toc.yml";
    let latestAsyncToc = fixToc(asyncTocPath, "../api-extractor-inputs-excel-async/excel.d.ts");
    fsx.writeFileSync(asyncTocPath, jsyaml.safeDump(latestAsyncToc));

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

    let origToc = (jsyaml.safeLoad(fsx.readFileSync(tocPath).toString()) as Toc);
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

    // create folders for Excel subcategories
    let excelEnumFilter = generateEnumList(fsx.readFileSync(sourceDtsPath).toString());
    let excelFilter: string[] = ["Interfaces"].concat(excelEnumFilter);

    // process all packages except 'office' (Common "Shared" API)
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            // fix host capitalization
            let packageName = (packageItem.name.substr(0, 1).toUpperCase() + packageItem.name.substr(1)).replace(/\-/g, ' ');
            membersToMove.items = packageItem.items;

            if (packageName.toLocaleLowerCase().includes('excel')) {
                let enumList = membersToMove.items.filter(item => {
                    return excelEnumFilter.indexOf(item.name) >= 0;
                });
                let primaryList = membersToMove.items.filter(item => {
                    return excelFilter.indexOf(item.name) < 0;
                });

                let excelEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
                primaryList.unshift(excelEnumRoot);
                newToc.items[0].items.push({
                    "name": packageName,
                    "uid": packageItem.uid,
                    "items": primaryList as any
                });
            }
        });
    });

    return newToc;
}
