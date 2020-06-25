IF EXIST "node_modules" (
    rmdir "node_modules" /s /q
)

IF EXIST "scripts\node_modules" (
    rmdir "scripts\node_modules" /s /q
)

IF EXIST "json" (
    rmdir "json" /s /q
)

call md json

IF EXIST "yaml" (
    rmdir "yaml" /s /q
)

call md yaml

call npm install

pushd scripts
call npm install
call npm run build
call node preprocessor.js
popd

pushd api-extractor-inputs-excel
call ..\node_modules\.bin\api-extractor run
popd
pushd api-extractor-inputs-excel-async
call ..\node_modules\.bin\api-extractor run
popd

call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel --output-folder .\yaml\excel --office
call .\node_modules\.bin\api-documenter yaml --input-folder .\json\excel-async --output-folder .\yaml\excel-async --office

pushd scripts
call node postprocessor.js
popd

pause