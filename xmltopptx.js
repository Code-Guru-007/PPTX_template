const fs = require('fs');
const AdmZip = require('adm-zip');

const directoryPath = './outputxml';

fs.promises.mkdir('./hello/hello')
const zip = new AdmZip();

// Read all files in the directory
const files = fs.readdirSync(directoryPath);

// Loop through each file in the directory
for (const file of files) {
  // Do something with the file
  if(file.includes('image')) zip.addFile(`ppt/media/${file}`, fs.readFileSync(`./outputxml/${file}`));
  else if(file.includes('rels')) 
    {
        if(file.includes('slide')) zip.addFile(`ppt/slides/_rels/${file}`, fs.readFileSync(`./outputxml/${file}`));
        else if(file.includes('presentation')) zip.addFile(`ppt/_rels/${file}`, fs.readFileSync(`./outputxml/${file}`));
        else if(file.includes('handoutMaster')) zip.addFile(`ppt/handoutMasters/_rels/${file}`, fs.readFileSync(`./outputxml/${file}`));
        else zip.addFile(`_rels/${file}`, fs.readFileSync(`./outputxml/${file}`));
    }
  else if(file==='presentation.xml') zip.addFile(`ppt/${file}`, fs.readFileSync(`./outputxml/${file}`));
  else if(file.includes('core') || file.includes('app') || file.includes('custom')) zip.addFile(`docprops/${file}`, fs.readFileSync(`./outputxml/${file}`));
  else if(file.includes('slide')) zip.addFile(`ppt/slides/${file}`, fs.readFileSync(`./outputxml/${file}`));
  else if(file.includes('[Content_Types]')) zip.addFile(`${file}`, fs.readFileSync(`./outputxml/${file}`));
  else  zip.addFile(`ppt/${file}`, fs.readFileSync(`./outputxml/${file}`));

}


// Save the zip archive as a PowerPoint file
fs.writeFileSync('./outputxml/file.pptx', zip.toBuffer());




// Specify the directory path
