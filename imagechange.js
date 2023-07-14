const unzipper = require("unzipper");
const fs = require("fs");
const path = require('path');
const xpath = require('xpath');
const dom = require('xmldom').DOMParser;
const fsextra = require('fs-extra');
const archiver = require('archiver');

const layouturl = "./pptx/ppt/slideLayouts"
const slideurl = "./pptx/ppt/slides"

const data = [
    {
        "1.1" : {
            "type" : "text",
            "value" : "This is title",
            "layout" : "false"
        },
        "1.2" : {
            "type" : "table",
            "value" : [
                { "name": "John", "phone": "+33653454343" },
                { "name": "John", "phone": "+33653454343" }
            ],
            "layout" : "false"
        },
        "1.3" : {
            "type" : "image",
            "value" : "./image/image1.jpg",
            "alt" : "{Alt_Page1_1.1}",
            "layout" : "true"
        },
        "1.4" : {
            "type" : "text",
            "value" : "This is subtext",
            "layout" : "false"
        }
    },
    {
        "2.1" : {
            "type" : "text",
            "value" : "This is title",
            "layout" : "false"
        },
        "2.2" : {
            "type" : "table",
            "value" : [
                { "name": "John", "phone": "+33653454343" },
                { "name": "John", "phone": "+33653454343" }
            ],
            "layout" : "false"
        },
        "2.3" : {
            "type" : "image",
            "value" : "./image/image.jpg",
            "alt" : "{Alt_Image1}",
            "layout" : "false"
        },
        "2.4" : {
            "type" : "text",
            "value" : "This is subtext",
            "layout" : "false"
        }
    }
  ]

  const findImageIDXml = (location, n, alt) => {

    console.log(`-------${location}/${n}`)
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
        console.log('Image copied successfully.');
    });
    });
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

// fs.createReadStream("./textchange.pptx")
//   .pipe(unzipper.Parse())
//   .on("entry", (entry) => {
//     const fileName = entry.path;
//     folderpath="./pptx/"
//     try {
//         fs.mkdirSync(folderpath, { recursive: true });
//         // console.log(`Folder created at ${path}`);
//         } catch (err) {
//         if (err.code !== 'EEXIST') {
//         //   console.error(`Error creating folder at ${path}:`, err);
//         } else {
//         //   console.log(`Folder already exists at ${path}`);
//         }
//         }
//     if (fileName.includes("/")) {
//         splitfilename = fileName.split('/')
//         for(let i=0;i<splitfilename.length-1;i++){
//             folderpath = folderpath + splitfilename[i] +"/";
//             // console.log(path)
//             try {
//                 fs.mkdirSync(folderpath, { recursive: true });
//                 // console.log(`Folder created at ${path}`);
//               } catch (err) {
//                 if (err.code !== 'EEXIST') {
//                 //   console.error(`Error creating folder at ${path}:`, err);
//                 } else {
//                 //   console.log(`Folder already exists at ${path}`);
//                 }
//               }
//         }
//         entry.pipe(fs.createWriteStream(`./pptx/${fileName}`));
//     }
//     else entry.pipe(fs.createWriteStream(`./pptx/${fileName}`));
//   })
//   .on('end', ()=> {
//     imageChange()
//     zipToPptx()
//   });



function imageChange(){
    for(pagenum in data) {
        const pagedata = data[pagenum]
        const filenum = parseInt(pagenum)+1
        for(key in pagedata) ///{'key':{}}
        {
            const changedata = pagedata[key]
            if(changedata["type"] == "image")
            {
                console.log(key, "---> ", changedata['layout'] )
                if(changedata['layout'] == "false"){
                    
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

imageChange()

function zipToPptx(){

    const outputZipPath = './output.pptx';
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

        const folderPath = './pptx';

        fsextra.remove(folderPath)
        .then(() => {
            console.log('Folder deleted successfully.');
        })
        .catch((err) => {
            console.error('Error deleting folder:', err);
        });

    });

}






