const xpath = require('xpath');
const { DOMParser } = require('xmldom');
const fs = require('fs');

// Read the XML file
const xml = fs.readFileSync('./pptx/ppt/slides/_rels/slide2.xml.rels', 'utf8');

const doc = new DOMParser().parseFromString(xml);

const nsResolver = (prefix) => {
  if (prefix === '') {
    return 'http://schemas.openxmlformats.org/package/2006/relationships';
  }
  return null;
};

const targetXPathExpression = '/*/*[@Id="rId2"]/@Target';
const selectedTargets = xpath.select(targetXPathExpression, doc, nsResolver);

const targetValue = selectedTargets.nodeValue;
console.log(targetValue);