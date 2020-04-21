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
    // delete everything except the overview file from the /docs folder
    fsx.readdirSync(docsDestination)
        .filter(filename => filename !== "overview.md")
        .forEach(filename => fsx.removeSync(docsDestination + '/' + filename));

    console.log(`Copying docs output files to: ${docsDestination}`);
    // copy docs output to /docs/docs-ref-autogen folder
    fsx.readdirSync(docsSource)
        .forEach(filename => {
        fsx.copySync(
            docsSource + '/' + filename,
            docsDestination + '/' + filename
        );
    });

    // fix all the individual TOC files
    console.log("Writing TOC for excel");
    const tocPath = path.resolve(`${docsDestination}/excel`) + "/toc.yml";
    let latestToc = fixToc(tocPath, "../api-extractor-inputs-excel/excel.d.ts");
    fsx.writeFileSync(tocPath, jsyaml.safeDump(latestToc));

    console.log("Writing TOC for excel");
    const asyncTocPath = path.resolve(`${docsDestination}/excel-async`) + "/toc.yml";
    let latestAsyncToc = fixToc(asyncTocPath, "../api-extractor-inputs-excel-async/excel.d.ts");
    fsx.writeFileSync(tocPath, jsyaml.safeDump(latestAsyncToc));

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
        "href": "../overview.md"
    }] as any;

    // create folders for Excel subcategories
    let excelEnumFilter = generateEnumList(fsx.readFileSync(sourceDtsPath).toString());
    let excelIconSetFilter : string [] = ["FiveArrowsGraySet", "FiveArrowsSet", "FiveBoxesSet", "FiveQuartersSet", "FiveRatingSet", "FourArrowsGraySet", "FourArrowsSet", "FourRatingSet", "FourRedToBlackSet", "FourTrafficLightsSet", "IconCollections", "ThreeArrowsGraySet", "ThreeArrowsSet", "ThreeFlagsSet",  "ThreeSignsSet", "ThreeStarsSet",  "ThreeSymbols2Set", "ThreeSymbolsSet", "ThreeTrafficLights1Set", "ThreeTrafficLights2Set", "ThreeTrianglesSet"];
    let excelFilter: string[] = ["Interfaces"];
    excelFilter = excelFilter.concat(excelEnumFilter).concat(excelIconSetFilter);

    // process all packages except 'office' (Common "Shared" API)
    origToc.items.forEach((rootItem, rootIndex) => {
        rootItem.items.forEach((packageItem, packageIndex) => {
            if (packageItem.name !== 'office') {
                // fix host capitalization
                let packageName = (packageItem.name.substr(0, 1).toUpperCase() + packageItem.name.substr(1)).replace(/\-/g, ' ');
                membersToMove.items = packageItem.items;

                if (packageName.toLocaleLowerCase().includes('excel')) {
                    let enumList = membersToMove.items.filter(item => {
                            return excelEnumFilter.indexOf(item.name) >= 0;
                        });
                    let iconSetList = membersToMove.items.filter(item => {
                            return excelIconSetFilter.indexOf(item.name) >= 0;
                    });
                    let primaryList = membersToMove.items.filter(item => {
                        return excelFilter.indexOf(item.name) < 0;
                    });

                    let excelEnumRoot = {"name": "Enums", "uid": "", "items": enumList};
                    let excelIconSetRoot = {"name": "Icon Sets", "uid": "", "items": iconSetList};
                    primaryList.unshift(excelIconSetRoot);
                    primaryList.unshift(excelEnumRoot);
                    newToc.items[0].items.push({
                        "name": packageName,
                        "uid": packageItem.uid,
                        "items": primaryList as any
                    });
                } else {
                    if (membersToMove.items) {
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items": membersToMove.items as any
                        });
                    } else {
                        newToc.items[0].items.push({
                            "name": packageName,
                            "uid": packageItem.uid,
                            "items": [] as any
                        });
                    }
                }
            }
        });
    });

    return newToc;
}
