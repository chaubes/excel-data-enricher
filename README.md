# excel-data-enricher

`npm install`

Usage: `node index.js <INPUT_DIR_PATH>`

_Example: node index.js test_

### Files In the INPUT_DIR_PATH
* config.properties - contains two property keys:
  * header_fields: comma-separated header fields required in the output file
  * input_files: comma-separated names of inout xlsx files
* input xlsx files - The comma separated files specified in the config.properties should be present in the input directory

### Output Excel File
Output excel file gets created at the path `<INPUT_DIR_PATH>/OutputFile.xlsx`