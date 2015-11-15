/**
 * To run the script, use:
 *
 * node index.js
 *
 * The script expects two files:
 *
 * * input\TypeScript Language Specification.docx
 * * input\spec.md
 *
 * and the final spec.md is produced to "output\spec.md" along with image files.
 */

var yauzl = require("yauzl");
var fs = require("fs");
var path = require("path");
var imagesFileNames = [];

yauzl.open("./input/TypeScript Language Specification.docx", function (err, zipfile) {
  if (err) {
    throw err;
  }

  zipfile.on("entry", function (entry) {
    if (/word\/media\/image[0-9]+.png$/i.test(entry.fileName)) {
      console.log(entry.fileName);

      var imageFileName = path.basename(entry.fileName);
      imagesFileNames.push(imageFileName);

      zipfile.openReadStream(entry, function (err, readStream) {
        if (err) {
          throw err;
        }

        readStream.pipe(fs.createWriteStream("./output/" + imageFileName));
      });
    }
  }).on("close", function () {
    var contents = fs.readFileSync('./input/spec.md', 'utf8');

    // Sort files
    imagesFileNames.sort(function (a, b) {
      return a.localeCompare(b);
    });

    var newContents = contents.replace(/^\/$/mg, function (match) {
      return "[Image](" + imagesFileNames.shift() + ")";
    });

    // Write output
    fs.writeFileSync('./output/spec.md', newContents, 'utf8');
    console.log("Done.");
  });
});