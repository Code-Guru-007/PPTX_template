const unzipper = require("unzipper");
const fs = require("fs");

fs.createReadStream("masterslide.pptx")
  .pipe(unzipper.Parse())
  .on("entry", (entry) => {
    const fileName = entry.path;
    console.log(fileName)
    // if (fileName.includes("ppt/slides/_rels")) {
    //   // Extract the master slide layout XML
    //   // entry.pipe(fs.createWriteStream(`./outputxml/${fileName.split('/')[3]}`));
    // } 
    // else {
    //   entry.autodrain();
    // }
  });
