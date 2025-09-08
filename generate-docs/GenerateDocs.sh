#!/bin/sh
if [ -d "node_modules" ]; then
    rm -rf "node_modules"
fi

if [ -d "scripts/node_modules" ]; then
    rm -rf "scripts/node_modules"
fi

if [ -d "json" ]; then
    rm -rf "json"
fi

mkdir -p json/json-preview
mkdir -p json/json-release

if [ -d "yaml" ]; then
    rm -rf "yaml"
fi

mkdir -p yaml/yaml-preview
mkdir -p yaml/yaml-release

npm install

pushd scripts
npm install
npm run build
node preprocessor.js
popd

# Create release versions by removing @beta APIs
./node_modules/.bin/version-remover api-extractor-inputs-excelscript/excelscript.d.ts api-extractor-inputs-excelscript-release/excelscript.d.ts "@beta"
./node_modules/.bin/version-remover api-extractor-inputs-officescript/officescript.d.ts api-extractor-inputs-officescript-release/officescript.d.ts "@beta"

# Generate what's new reports comparing preview vs release
./node_modules/.bin/whats-new api-extractor-inputs-excelscript/excelscript.d.ts api-extractor-inputs-excelscript-release/excelscript.d.ts ../docs/includes/excelscript-preview javascript/api/office-scripts/excelscript. office-scripts-link-config.json
./node_modules/.bin/whats-new api-extractor-inputs-officescript/officescript.d.ts api-extractor-inputs-officescript-release/officescript.d.ts ../docs/includes/officescript-preview javascript/api/office-scripts/officescript. office-scripts-link-config.json

pushd api-extractor-inputs-excelscript
../node_modules/.bin/api-extractor run
popd

pushd api-extractor-inputs-officescript
../node_modules/.bin/api-extractor run
popd

pushd api-extractor-inputs-excelscript-release
../node_modules/.bin/api-extractor run
popd

pushd api-extractor-inputs-officescript-release
../node_modules/.bin/api-extractor run
popd

./node_modules/.bin/api-documenter yaml --input-folder ./json/json-preview --output-folder ./yaml/yaml-preview --office
./node_modules/.bin/api-documenter yaml --input-folder ./json/json-release --output-folder ./yaml/yaml-release --office

pushd scripts
node postprocessor.js
popd

./node_modules/.bin/reference-coverage-tester reference-coverage-tester.json

wait
