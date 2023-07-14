const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const path = require("path");
const fs = require("fs");
const xpath = require('xpath');
const dom = require('xmldom').DOMParser;
const fsextra = require('fs-extra');
const archiver = require('archiver');
const AdmZip = require('adm-zip');

const data = require('./data.json')

let render_json = {}
const layouturl = "./pptx/ppt/slideLayouts"
const slideurl = "./pptx/ppt/slides"



const content = fs.readFileSync(
  path.resolve(__dirname, "input.pptx"),
  "binary"
);

const zip = new PizZip(content);
const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
});

for(pagenum in data) { ///  [{},{},{}]
  const pagedata = data[pagenum]
  for(key in pagedata) ///{'key':{}}
  {
    const changedata = pagedata[key]
    if(changedata["type"] == "text" || changedata["type"] == "table")
      {
        render_json[key] = changedata['value']
      }

  }
}
doc.render(render_json);

const buf = doc.getZip().generate({
  type: "nodebuffer",
  compression: "DEFLATE",
});
fs.writeFileSync(path.resolve(__dirname, "textchange.pptx"), buf);

const findImageIDXml = (location, n, alt) => {

  const xmlData = fs.readFileSync(`${location}/${n}`, 'utf8');
  const doc = new dom().parseFromString(xmlData);

  const namespaces = {
      p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
      a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    };
    
    // Create the namespace resolver
    const nsResolver = (prefix) => namespaces[prefix] || null;
    
    // Select the desired <pic> tag using XPath with registered namespace
    const xpathExpression = `//*[local-name()='pic'][*[local-name()='nvPicPr']/*[local-name()='cNvPr'][@descr='${alt}']]`;
    const selectedPicTag  = xpath.select(xpathExpression, doc, nsResolver);
    
    if (selectedPicTag) {
      const blipXPathExpression = `.//*[local-name()='blip' and namespace-uri()='${namespaces.a}']`;
      const selectedBlipElements = xpath.select(blipXPathExpression, selectedPicTag, nsResolver);
    
      if (selectedBlipElements) {
        const rembed = selectedBlipElements.getAttribute('r:embed');
        return rembed;
      }
    } else {
      return false;
    }
}

const findImageUrlXml = (location, n, id) => {
  const xml = fs.readFileSync(`${location}/_rels/${n}.rels`, 'utf8');

  const doc = new dom().parseFromString(xml);

  const nsResolver = (prefix) => {
  if (prefix === '') {
      return 'http://schemas.openxmlformats.org/package/2006/relationships';
  }
  return null;
  };

  const targetXPathExpression = `/*/*[@Id="${id}"]/@Target`;
  const selectedTargets = xpath.select(targetXPathExpression, doc, nsResolver);

  const targetValue = selectedTargets.nodeValue;
  return targetValue;
}
const findImageLayout = (pagenum) => {
  const xml = fs.readFileSync(`./pptx/ppt/slides/_rels/slide${pagenum}.xml.rels`, 'utf8');

  const doc = new dom().parseFromString(xml);

  const nsResolver = (prefix) => {
  if (prefix === '') {
      return 'http://schemas.openxmlformats.org/package/2006/relationships';
  }
  return null;
  };

  const targetXPathExpression = `/*/*[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"]/@Target`;
  const selectedTargets = xpath.select(targetXPathExpression, doc, nsResolver);

  const targetValue = selectedTargets.nodeValue.split('/')[2];
  return targetValue;
}

const replaceImage = (targetImage, sourceImagePath) => {
  const destinationDirectory = './pptx/ppt/media/';
  const destinationImagePath = path.join(destinationDirectory, targetImage);

  // Delete the existing image16.png
  fs.unlink(destinationImagePath, (err) => {
  if (err) {
      console.error('Error deleting the existing image:', err);
  }

  // Copy the new image to the destination directory
  fs.copyFile(sourceImagePath, destinationImagePath, (err) => {
      if (err) {
      console.error('Error copying the new image:', err);
      }
  });
  });
}



// Specify the path to the .pptx file
const pptxFilePath = 'textchange.pptx';

// Create an instance of AdmZip
const unzippptx = new AdmZip(pptxFilePath);

// Extract the contents of the .pptx file to the specified directory
unzippptx.extractAllTo('./pptx', true);

const imageChange = () => {
  for(pagenum in data) {
      const pagedata = data[pagenum]
      const filenum = parseInt(pagenum)+1
      for(key in pagedata) ///{'key':{}}
      {
          const changedata = pagedata[key]
          if(changedata["type"] == "image")
          {
            if(changedata['masterslide'] == "false"){                  
                  try {
                      imageid = findImageIDXml(slideurl, "slide"+filenum +'.xml', changedata['alt'])
                      imageUrl = findImageUrlXml(slideurl, "slide"+filenum + '.xml', imageid)
                      imageName = imageUrl.split('/')[2]
                      if(imageName){
                          replaceImage(imageName, changedata['value'])
                      }
                  } catch (error) {
                      
                  }
              }
              else {
                  try {
                      imagelayout = findImageLayout(parseInt(pagenum)+1)
                      imageid = findImageIDXml(layouturl, imagelayout, changedata['alt'])
                      imageUrl = findImageUrlXml(layouturl, imagelayout, imageid)
                      imageName = imageUrl.split('/')[2]
                      if(imageName){
                          replaceImage(imageName, changedata['value'])
                      }
                  } catch (error) {
                      
                  }
              }
          }

      }
  }
}

const zipToPptx = () => {
  const outputZipPath = './output.pptx';
  const inputFolderPath = './pptx';
  const outputZipStream = fs.createWriteStream(outputZipPath);
  const archive = archiver('zip', { zlib: { level: 9 } });
  archive.on('error', (err) => {
  console.error('Error creating the zip file:', err);
  });
  archive.pipe(outputZipStream);
  archive.directory(inputFolderPath, false);
  archive.finalize();
  outputZipStream.on('close', () => {
      console.log('Zip file created successfully.');
      const folderPath = './pptx';
      fsextra.remove(folderPath)
      .then(() => {
      })
      .catch((err) => {
          console.error('Error deleting folder:', err);
      });
      fs.unlink("./textchange.pptx", (err) => {
        if (err) {
          return;
        }
      });
  });

}

imageChange()
zipToPptx()




