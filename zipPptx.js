const fs = require('fs');
const archiver = require('archiver');

const outputZipPath = './pptx.pptx';
const inputFolderPath = './pptx';

// Create a writable stream for the output zip file
const outputZipStream = fs.createWriteStream(outputZipPath);

// Create a new archive instance
const archive = archiver('zip', { zlib: { level: 9 } });

// Event listener for the 'error' event of the archive
archive.on('error', (err) => {
  console.error('Error creating the zip file:', err);
});

// Pipe the archive data to the output stream
archive.pipe(outputZipStream);

// Add the contents of the "pptx" folder to the archive
archive.directory(inputFolderPath, false);

// Finalize the archive
archive.finalize();

// Event listener for the 'close' event of the output stream
outputZipStream.on('close', () => {
  console.log('Zip file created successfully.');
});
