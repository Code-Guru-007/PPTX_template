const xpath = require('xpath');
const { DOMParser } = require('xmldom');
const fs = require('fs');

// Read the XML file
const xml = fs.readFileSync('./pptx/ppt/slides/slide2.xml', 'utf8');

// Parse the XML document
const doc = new DOMParser().parseFromString(xml);

// Register the namespace prefixes and URIs
const namespaces = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
};

// Create the namespace resolver
const nsResolver = (prefix) => namespaces[prefix] || null;

// Select the desired <pic> tag using XPath with registered namespace
const altImage = '{Alt_Image1}';
const xpathExpression = `//*[local-name()='pic'][*[local-name()='nvPicPr']/*[local-name()='cNvPr'][@descr='${altImage}']]`;
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
