#!/usr/bin/env node --harmony
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const fsx = require("fs-extra");
const jsyaml = require("js-yaml");
const path = require("path");
var DescriptionRating;
(function (DescriptionRating) {
    DescriptionRating["Missing"] = "Missing";
    DescriptionRating["Poor"] = "Poor";
    DescriptionRating["Fine"] = "Fine";
    DescriptionRating["Good"] = "Good";
})(DescriptionRating || (DescriptionRating = {}));
class CoverageRating {
}
class ClassCoverageRating {
    constructor() {
        this.apiRatings = new Map();
        this.classRating = {
            descriptionRating: DescriptionRating.Missing,
            hasExample: false
        };
    }
}
class ApiFieldYaml {
}
class ApiPropertyYaml {
}
class ApiMethodYaml {
}
let ratingMap = new Map();
const docsSource = path.resolve("../../docs/docs-ref-autogen/excel/excelscript");
fsx.readdirSync(docsSource)
    .forEach(filename => {
    console.log(`Checking ${filename}.`);
    let ymlFile = jsyaml.load(fsx.readFileSync(docsSource + '/' + filename).toString());
    let rating = rateClass(ymlFile);
    ratingMap.set(ymlFile.name, rating);
});
let csvString = convertToCsv(ratingMap);
fsx.writeFileSync(path.resolve("./") + "/API Coverage Report.csv", csvString);
process.exit(0);
function rateClass(classYml) {
    var _a, _b, _c;
    let ymlCoverage = new ClassCoverageRating();
    ymlCoverage.classRating = rateClassDescription(classYml);
    (_a = classYml.fields) === null || _a === void 0 ? void 0 : _a.forEach((field) => {
        // Note: examples in enum fields are intentionally not supported.
        ymlCoverage.apiRatings.set(field.name, {
            descriptionRating: rateDescriptionString(field.summary),
            hasExample: false
        });
    });
    (_b = classYml.properties) === null || _b === void 0 ? void 0 : _b.forEach((field) => {
        ymlCoverage.apiRatings.set(field.name, rateFieldDescription(field));
    });
    (_c = classYml.methods) === null || _c === void 0 ? void 0 : _c.forEach((field) => {
        let name = field.name.indexOf(",") < 0 ? field.name : field.name.substring(0, field.name.indexOf(","));
        ymlCoverage.apiRatings.set(name, rateFieldDescription(field));
    });
    return ymlCoverage;
}
function rateClassDescription(classYml) {
    var _a;
    let rating;
    let indexOfExample = (_a = classYml.remarks) === null || _a === void 0 ? void 0 : _a.indexOf("#### Examples");
    if (indexOfExample > 0) {
        rating = {
            descriptionRating: rateDescriptionString((classYml.summary + " " + classYml.remarks.substring(0, indexOfExample)).trim()),
            hasExample: true
        };
    }
    else {
        rating = {
            descriptionRating: rateDescriptionString((classYml.summary + " " + classYml.remarks).trim()),
            hasExample: false
        };
    }
    return rating;
}
function rateFieldDescription(fieldYml) {
    var _a, _b;
    let rating;
    let indexOfExample = (_a = fieldYml.syntax.return.description) === null || _a === void 0 ? void 0 : _a.indexOf("#### Examples");
    if (indexOfExample > 0) {
        rating = {
            descriptionRating: rateDescriptionString((fieldYml.summary + " " + fieldYml.remarks.substring(0, indexOfExample)).trim()),
            hasExample: true
        };
    }
    else {
        rating = {
            descriptionRating: rateDescriptionString((fieldYml.summary + " " + fieldYml.remarks).trim()),
            hasExample: false
        };
    }
    if (fieldYml instanceof ApiMethodYaml) {
        let methodYml = fieldYml;
        let descriptionRatings = [rateDescriptionString(methodYml.syntax.return.description), rating.descriptionRating];
        (_b = methodYml.syntax.parameters) === null || _b === void 0 ? void 0 : _b.forEach((parameter) => {
            descriptionRatings.push(rateDescriptionString(parameter.description));
        });
        rating.descriptionRating = averageDescriptionRatings(descriptionRatings);
    }
    return rating;
}
function rateDescriptionString(description) {
    if (description === "") {
        return DescriptionRating.Missing;
    }
    let sentenceCount = description.split(". ").length;
    let wordCount = description.split(" ").length;
    if (wordCount < 10) {
        return DescriptionRating.Poor;
    }
    else if (sentenceCount < 2) {
        return DescriptionRating.Fine;
    }
    else {
        return DescriptionRating.Good;
    }
}
function averageDescriptionRatings(ratings) {
    let ratingScore = 0;
    ratings.forEach((rating) => {
        switch (rating) {
            case DescriptionRating.Missing:
                return DescriptionRating.Missing;
            case DescriptionRating.Poor:
                ratingScore += 1;
                break;
            case DescriptionRating.Fine:
                ratingScore += 2;
                break;
            case DescriptionRating.Good:
                ratingScore += 3;
                break;
        }
    });
    ratingScore /= ratings.length;
    if (ratingScore === 3) {
        return DescriptionRating.Good;
    }
    else if (ratingScore > 2) {
        return DescriptionRating.Fine;
    }
    else {
        return DescriptionRating.Poor;
    }
}
function convertToCsv(apiCoverage) {
    let csvString = "Class,Field,Description Rating, Has Example?\n";
    apiCoverage.forEach((coverage, className) => {
        csvString += `${className},,${coverage.classRating.descriptionRating},${coverage.classRating.hasExample}\n`;
        coverage.apiRatings.forEach((fieldCoverage, fieldName) => {
            csvString += `,${fieldName},${fieldCoverage.descriptionRating},${fieldCoverage.hasExample}\n`;
        });
    });
    return csvString;
}
//# sourceMappingURL=coverage-tester.js.map