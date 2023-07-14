const fs = require('fs');
const archiver = require('archiver');

// Specify the path to the XML file
const xmlFilePath = './file.xml';

// Specify the path to the output PowerPoint file
const pptxFilePath = './file.pptx';

// Create a new archive
const output = fs.createWriteStream(pptxFilePath);
const archive = archiver('zip', { zlib: { level: 9 } });

// Pipe the output to the archive
archive.pipe(output);

// Add the necessary files and folders to the archive
archive.directory('path/to/pptx', 'pptx');
archive.file(xmlFilePath, { name: 'pptx/ppt/slides/slide1.xml' });

// Finalize the archive
archive.finalize();