const fs = require("fs");
const { parseString } = require("xml2js");
fs.readFile("masterSlideLayout.xml", "utf-8", (err, data) => {
    if (err) {
      console.error("Error reading XML file:", err);
      return;
    }
  
    // Convert XML to JSON
    parseString(data, (err, result) => {
      if (err) {
        console.error("Error parsing XML:", err);
        return;
      }
  
      // JSON output
      const jsonOutput = JSON.stringify(result, null, 2);
      fs.writeFileSync("masterslidelayout.json", jsonOutput);
  
      // Save JSON to a file if needed
      // fs.writeFileSync("masterslidelayout.json", jsonOutput);
    });
  });
  