const officegen = require("officegen");
const fs = require("fs");
const path = require("path");

exports.document = (req, res, next) => {
  let docx = officegen("docx");
  let firstName = "Rohit";
  let lastName = "Thakur";
  docx.on("finalize", function (written) {
    console.log("Finish to create a Microsoft Word document.");
    res.send("Done");
    res.end();
  });
  docx.on("error", function (err) {
    console.log(err);
  });

  let pObj = docx.createP();
  pobj = docx.createP();
  pObj.addText(` Hello ${firstName} ${lastName}`, {
    font_face: "Arial",
    font_size: 40,
  });

  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText(
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Proin libero nunc consequat interdum varius sit amet. Quis hendrerit dolor magna eget est lorem. At erat pellentesque adipiscing commodo elit at imperdiet dui accumsan. Orci dapibus ultrices in iaculis nunc sed augue lacus viverra. Et malesuada fames ac turpis egestas integer. Suspendisse ultrices gravida dictum fusce ut placerat orci. In arcu cursus euismod quis viverra nibh cras pulvinar. Mattis vulputate enim nulla aliquet porttitor lacus luctus accumsan tortor. Sit amet mattis vulputate enim nulla aliquet porttitor lacus. Ultrices sagittis orci a scelerisque purus semper eget. Dui accumsan sit amet nulla facilisi morbi tempus. A diam maecenas sed enim ut sem. In metus vulputate eu scelerisque felis imperdiet proin fermentum. Leo urna molestie at elementum eu facilisis. Feugiat vivamus at augue eget arcu dictum varius duis."
  );

  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText(
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Proin libero nunc consequat interdum varius sit amet. Quis hendrerit dolor magna eget est lorem. At erat pellentesque adipiscing commodo elit at imperdiet dui accumsan. Orci dapibus ultrices in iaculis nunc sed augue lacus viverra. Et malesuada fames ac turpis egestas integer. Suspendisse ultrices gravida dictum fusce ut placerat orci. In arcu cursus euismod quis viverra nibh cras pulvinar. Mattis vulputate enim nulla aliquet porttitor lacus luctus accumsan tortor. Sit amet mattis vulputate enim nulla aliquet porttitor lacus. Ultrices sagittis orci a scelerisque purus semper eget. Dui accumsan sit amet nulla facilisi morbi tempus. A diam maecenas sed enim ut sem. In metus vulputate eu scelerisque felis imperdiet proin fermentum. Leo urna molestie at elementum eu facilisis. Feugiat vivamus at augue eget arcu dictum varius duis."
  );

  docx.putPageBreak();

  var table = [
    [
      {
        val: "No.",
        opts: {
          cellColWidth: 100,
          b: true,
          sz: "20",
          spacingBefore: 10,
          spacingAfter: 10,
          spacingLine: 20,
          spacingLineRule: "atLeast",
          shd: {
            fill: "7F7F7F",
            themeFill: "text1",
            themeFillTint: "80",
          },
          fontFamily: "Avenir Book",
        },
      },
      {
        val: "Title1",
        opts: {
          b: true,
          color: "A00000",
          align: "right",
          shd: {
            fill: "92CDDC",
            themeFill: "text1",
            themeFillTint: "80",
          },
        },
      },
    ],
    [1, "Col 1"],
    [2, "Col2"],
    [3, "Col 3"],
    [4, "Col 4"],
    [5, "Col 5"],
    [6, "Col 6"],
  ];

  var tableStyle = {
    tableColWidth: 100,
    tableSize: 20,
    tableAlign: "left",
    tableFontFamily: "Comic Sans MS",
    borders: true,
    borderSize: 2,
    columns: [{ width: 1 }, { width: 1 }, { width: 1 }], // Table logical columns
  };

  docx.createTable(table, tableStyle);
  let out = fs.createWriteStream("example.docx");

  out.on("error", function (err) {
    console.log(err);
  });

  docx.generate(out);
};
