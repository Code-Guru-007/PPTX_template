const fsextra = require('fs-extra');

const folderPath = './pptx';

// Delete the folder and its contents
fsextra.remove(folderPath)
  .then(() => {
    console.log('Folder deleted successfully.');
  })
  .catch((err) => {
    console.error('Error deleting folder:', err);
  });
