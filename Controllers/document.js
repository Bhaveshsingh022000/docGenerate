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
  // pobj = docx.createP();
  // pObj.addText(` Hello ${firstName} ${lastName}`, {
  //   font_face: "Arial",
  //   font_size: 40,
  // });
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();

  pobj = docx.createP();
  pObj.options.align = "center";
  pObj.addText(`Proposal:`, {
    font_face: "Times New Roman",
    font_size: 24,
    bold: true,
  });

  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText(
    "Spill Control and Countermeasure Plan (SPCC Plan) Site Evaluation & SPCC Drafting",
    {
      font_face: "Times New Roman",
      font_size: 18,
      bold: true,
    }
  );
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();

  pObj.addText("Submitted to:",{
    font_face: "Times New Roman",
      font_size: 14,
      italic: true,
  });
  pObj.addLineBreak();
  pObj.addLineBreak();

  const corpName= "Test Corp Name"
  const addressLine1 = "Test Address Line 1";
  const addressLine2 = "Test Address Line 2";

  pObj.addText(`${corpName}`,{
    font_face: "Times New Roman",
      font_size: 18,
      bold: true,
  });
  pObj.addLineBreak();

  pObj.addText(`${addressLine1}`,{
    font_face: "Times New Roman",
      font_size: 14,
      bold: true,
  });
  pObj.addLineBreak();

  pObj.addText(`${addressLine2}`,{
    font_face: "Times New Roman",
      font_size: 14,
      bold: true,
  });
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();

  const day = 21;
  const month = 12;
  const year = 2020;
  pObj.addText(`${day} ${month} ${year}`,{
    font_face: "Times New Roman",
      font_size: 18,
      bold: true,
  });
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();

  pObj.addImage(path.resolve(__dirname,"../public/images/logo.png"),{cx: 350, cy: 100});



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
