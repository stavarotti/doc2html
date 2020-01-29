const fs = require('fs');
const { JSDOM } = require('jsdom');
const mammoth = require('mammoth');
const pretty = require('pretty');
const inquirer = require('inquirer');
const minimist = require('minimist');
const opts = minimist(process.argv.slice(2));

/**
 * Script entry point.
 *
 * @private
 * @function main
 */
async function main() {
  try {
    // Prompt the user for the file name if not provided.
    if (!opts.file) {
      const filePath = await inquirer.prompt([
        {
          type: 'input',
          name: 'file',
          message: `Please enter the path to the Word document: `,
          validate: file => !!file
        }
      ]);

      opts.file = filePath.file;
    }

    // Prompt the user for the destination folder name
    if (!opts.dest) {
      const destFolderPath = await inquirer.prompt([
        {
          type: 'input',
          name: 'dest',
          message: `Please enter the destination folder: `,
          validate: dest => !!dest
        }
      ]);

      opts.dest = destFolderPath.dest;
    }

    // Parse the document and generate the files
    generateHTMLFiles(opts.file, opts.dest);
  } catch (e) {
    console.error(e);
  }
}

/**
 * Parses the recovered Microsoft Word file into HTML documents.
 *
 * @private
 * @function generateHTMLFiles
 *
 * @param {string} filePath The Word file path
 * @param {string} destFolderPath The destination folder
 */
async function generateHTMLFiles(filePath, destFolderPath) {
  const result = await mammoth.convertToHtml({ path: filePath });

  writeFiles(processConvertedDocument(result.value), destFolderPath);
}

/**
 * Converts the provided HTML into individual sections demarcated by header.
 *
 * @private
 * @function processConvertedDocument
 *
 * @param {string} html The converted document HTML.
 *
 * @returns {string[]} The HTML documents.
 */
function processConvertedDocument(html) {
  const HEADING_NODE_NAMES = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'];

  // Create a new DOM using the raw HTML result
  const { window } = new JSDOM(`<!DOCTYPE html><html><head></head><body>${html}</body></html>`);

  // Stores the HTML documents that'll be saved to disk
  let htmlDocs = [];

  // Captures the current document fragment (equivalent to a file)
  let htmlDoc = null;

  // Iterate over all direct children of the body element and generate the docs.
  [...window.document.querySelector('body').children].forEach((element, index) => {
    const elementNodeName = element.nodeName.toLowerCase();

    // Accounting for malformed or irregular documents that don't start with a heading
    if (index === 0 && HEADING_NODE_NAMES.indexOf(elementNodeName) === -1) {
      htmlDoc = {
        fragment: window.document.createDocumentFragment(),
        name: 'orphan'
      };

      htmlDoc.fragment.appendChild(element);
    }

    if (HEADING_NODE_NAMES.indexOf(elementNodeName) > -1) {
      // New document detected, so flush the current HTML Doc into the fileList
      if (htmlDoc) {
        htmlDocs.push(htmlDoc);
      }

      // Create a new reference
      htmlDoc = {
        fragment: window.document.createDocumentFragment(),
        name: element.textContent || ''
      };

      htmlDoc.fragment.appendChild(element);
    } else {
      // Didn't detect a header, so appending to the existing document
      htmlDoc.fragment.appendChild(element);
    }
  });

  return htmlDocs;
}

/**
 * Writes to disk, the HTML docs as individuals files.
 *
 * @private
 * @function writeFiles
 *
 * @param {Object[]} htmlDocs The HTML doc objects.
 * @param {string} destFolderPath The destination folder
 */
function writeFiles(htmlDocs = [], destFolderPath = './output') {
  const { window } = new JSDOM();

  if (!fs.existsSync(destFolderPath)) {
    fs.mkdirSync(destFolderPath);
  }

  htmlDocs.forEach(htmlDoc => {
    // Need a wrapper element in order to grab the fragment innerHTML
    const el = window.document.createElement('div');
    el.appendChild(htmlDoc.fragment);

    fs.writeFileSync(
      `${destFolderPath}/${htmlDoc.name.trim()}.html`,
      pretty(
        `<DOCTYPE html>
         <html>
           <head>
            <style>h1,h2,h3,h4,h5,h6 {font-size: 1rem; font-weight: bold}</style>
           </head>
           <body>${el.innerHTML}</body>
         </html>
        `
      )
    );
  });
}

// Start the process
main();
