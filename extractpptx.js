const AdmZip = require('adm-zip');

// Specify the path to the .pptx file
const pptxFilePath = 'textchange.pptx';

// Create an instance of AdmZip
const zip = new AdmZip(pptxFilePath);

// Extract the contents of the .pptx file to the specified directory
zip.extractAllTo('./path', true);