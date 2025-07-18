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

mkdir json

if [ -d "yaml" ]; then
    rm -rf "yaml"
fi

mkdir yaml

npm install

pushd scripts
npm install
npm run build
node preprocessor.js
popd

pushd api-extractor-inputs-excelscript
../node_modules/.bin/api-extractor run
popd

pushd api-extractor-inputs-officescript
../node_modules/.bin/api-extractor run
popd

./node_modules/.bin/api-documenter yaml --input-folder ./json --output-folder ./yaml --office

pushd scripts
node postprocessor.js
popd

./node_modules/.bin/reference-coverage-tester reference-coverage-tester.json

wait
