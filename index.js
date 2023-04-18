const propertiesReader = require('properties-reader');
const path = require('path');
const excelToJson = require('convert-excel-to-json');
const xlsx = require("json-as-xlsx");

const input_dir = process.argv[2];
const properties = propertiesReader(`${input_dir}${path.sep}config.properties`);
const outputFilePath = `${input_dir}${path.sep}OutputFile`;

let settings = {
    fileName: `${outputFilePath}`, // Name of the resulting spreadsheet
    extraLength: 3, // A bigger number means that columns will be wider
    writeMode: "writeFile", // The available parameters are 'WriteFile' and 'write'. This setting is optional. Useful in such cases https://docs.sheetjs.com/docs/solutions/output#example-remote-file
    writeOptions: {}, // Style options from https://docs.sheetjs.com/docs/api/write-options
    //RTL: true, // Display the columns from right-to-left (the default value is false)
};

let callback = function (sheet) {
    console.log(`Download complete....${outputFilePath}`);
};

const prepareOutputExcelData = (outputJsonArray, headerFieldsArray) => {
    let data = [
        {
            sheet: "OutputSheet",
            columns: headerFieldsArray.map(h => {
                return {
                    label: h, value: h
                };
            }),
            content: outputJsonArray,
        },
    ]
    return data;
}

const parseExcelToJson = (filePath) => {
    const data = [];
    const headerMap = new Map();
    const result = excelToJson({
        sourceFile: filePath,
        sheets: ['Sheet1'],
    });
    const dataArray = result['Sheet1']
    for(let i=0;i < dataArray.length; i++){
        const jsonData = {};
        Object.keys(dataArray[i]).forEach(h=> {
            if(i === 0) {
                headerMap.set(h, dataArray[i][h]);
            } else {
                jsonData[headerMap.get(h)] = dataArray[i][h];
            }
        });
        if(i > 0){
            data.push(jsonData);
        }
    }
    return data;
}

const prepareOutputJson = (allData, headerFieldsArray) => {
    const outputJsonArray = [];
    const headerSet = new Set(headerFieldsArray);

    allData.forEach(data => {
        const jsonData = {};
        Object.keys(data).forEach(k=> {
            if(headerSet.has(k)){
                jsonData[k] = data[k];
            }
        });

        const jsonDataKeys = new Set(Object.keys(jsonData));

        headerSet.forEach(h=> {
            if(!jsonDataKeys.has(h)){
                jsonData[h] = '';
            }
        });
        outputJsonArray.push(jsonData);
    });

    return outputJsonArray;
}

const init = async () => {
    const headerFields = properties.get('header_fields');
    const headerFieldsArray = headerFields.split(',');
    const inputFiles = properties.get('input_files');
    const inputFilesArray = inputFiles.split(',');
    const allFilesData = [];
    inputFilesArray.forEach(file=> {
        const filePath = `${input_dir}${path.sep}${file.trim()}`;
        const data = parseExcelToJson(filePath);
        allFilesData.push(...data);
    });
    const outputJsonArray = prepareOutputJson(allFilesData, headerFieldsArray);

    console.log(JSON.stringify(outputJsonArray, null, 2));

    // Write output file
    xlsx(prepareOutputExcelData(outputJsonArray, headerFieldsArray), settings, callback);
}

init();