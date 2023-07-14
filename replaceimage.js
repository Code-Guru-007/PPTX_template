
const fs = require('fs');
const path = require('path');

const sourceImagePath = './new/image.jpg';
const destinationDirectory = './pptx/ppt/media/';
const destinationImagePath = path.join(destinationDirectory, 'image16.png');

// Delete the existing image16.png
fs.unlink(destinationImagePath, (err) => {
  if (err) {
    console.error('Error deleting the existing image:', err);
  }

  // Copy the new image to the destination directory
  fs.copyFile(sourceImagePath, destinationImagePath, (err) => {
    if (err) {
      console.error('Error copying the new image:', err);
      return;
    }

    console.log('Image copied successfully.');
  });
});