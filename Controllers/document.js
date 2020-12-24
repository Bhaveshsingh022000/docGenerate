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

  pObj.addText("Submitted to:", {
    font_face: "Times New Roman",
    font_size: 14,
    italic: true,
  });
  pObj.addLineBreak();
  pObj.addLineBreak();

  const corpName = "Test Corp Name";
  const addressLine1 = "Test Address Line 1";
  const addressLine2 = "Test Address Line 2";

  pObj.addText(`${corpName}`, {
    font_face: "Times New Roman",
    font_size: 18,
    bold: true,
  });
  pObj.addLineBreak();

  pObj.addText(`${addressLine1}`, {
    font_face: "Times New Roman",
    font_size: 14,
    bold: true,
  });
  pObj.addLineBreak();

  pObj.addText(`${addressLine2}`, {
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
  pObj.addText(`${day} ${month} ${year}`, {
    font_face: "Times New Roman",
    font_size: 18,
    bold: true,
  });
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addLineBreak();

  pObj.addImage(path.resolve(__dirname, "../public/images/logo.png"), {
    cx: 350,
    cy: 100,
  });

  docx.putPageBreak();
  let pObj2 = docx.createP();
  pObj2.options.align = "center";

  var header = docx.getHeader().createP();
  header.addImage(path.resolve(__dirname, "../public/images/logo.png"), {
    cx: 280,
    cy: 80,
  });
  header.addText("Environmental Engineering and Consulting");

  pObj2.addText(`Pacific Engineering & Consulting`, {
    font_face: "Times New Roman",
    font_size: 14,
  });
  pObj2.addLineBreak();

  pObj2.addText(`Company Credentials and Certifications`, {
    font_face: "Times New Roman",
    font_size: 12,
    underline: true,
  });
  pObj2.addLineBreak();

  let pObj3 = docx.createP();
  pObj3.options.align = "left";
  pObj3.addText(`Name of Organization: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });

  pObj3.addText(`Pacific Engineering & Consulting `, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj3.addLineBreak();
  pObj3.addLineBreak();
  pObj3.addLineBreak();

  pObj3.addText(`Office Address: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });

  pObj3.addText(`1788 N Helm Ave Suite 112 Fresno CA, 93727 `, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj3.addLineBreak();
  pObj3.addLineBreak();
  pObj3.addLineBreak();

  pObj3.addText(`Web Address: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });

  pObj3.addText(`www.pacificmgt.com `, {
    font_face: "Times New Roman",
    font_size: 12,
    indentLeft: 1440,
  });
  pObj3.addLineBreak();
  pObj3.addLineBreak();
  pObj3.addLineBreak();

  pObj3.addText(`Telephone Number: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });

  pObj3.addText(`(559) 251-4060 `, {
    font_face: "Times New Roman",
    font_size: 12,
    indentLeft: 1440,
  });
  pObj3.addLineBreak();
  pObj3.addLineBreak();
  pObj3.addLineBreak();

  pObj3.addText(`Fax: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });

  pObj3.addText(`(559) 251-4060 `, {
    font_face: "Times New Roman",
    font_size: 12,
    indentLeft: 1440,
  });
  pObj3.addLineBreak();
  pObj3.addLineBreak();
  pObj3.addLineBreak();

  pObj3.addText(`DUNS Number: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });

  pObj3.addText(`196553770 `, {
    font_face: "Times New Roman",
    font_size: 12,
    indentLeft: 1440,
  });
  pObj3.addLineBreak();
  pObj3.addLineBreak();
  pObj3.addLineBreak();

  pObj3.addText(`Size of Company: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  2;

  pObj3.addText(
    `Small under NAICS 541330 - Engineering Services, 541620 - Environmental Consulting Service, 541690 - Other Scientific and Technical Consulting Services. `,
    {
      font_face: "Times New Roman",
      font_size: 12,
      indentLeft: 1440,
    }
  );
  pObj3.addLineBreak();
  pObj3.addLineBreak();
  pObj3.addLineBreak();

  pObj3.addText(`Point of Contact: `, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj3.addText(
    `Daniel Elliott Project Manager (559) 251-4060 ext. 105 daniel@pacificmgt.com`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );

  docx.putPageBreak();

  let pObj4 = docx.createListOfNumbers();
  pObj4.options.align = "left";
  pObj4.addText(`Introduction`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj4.addLineBreak();
  pObj4.addLineBreak();
  pObj4.addText(
    `Pacific Engineering & Consulting was founded in 1982 and is headquartered in Fresno, CA. We currently specialize in engineering and environmental services, specifically in performing certified inspection(s) and SPCC Plans. Pacific Engineering & Consulting has conducted Aboveground and Underground Storage Tank (AST/UST) inspections and testing for a variety of US Government agencies and Private Industry customers.  We have also done pressure testing, pressure vessel inspection, Spill Prevention Control and Countermeasure (SPCC) plans, Facility Response Plans (FRP). For government and commercial customers throughout the continental US. Pacific Engineering & Consulting is certified to conduct a variety of inspections including: storage tanks, piping, pressure vessels, OSHA’s Process Safety Management integrity inspections, Non-Destructive Examination, National Association of Corrosion Engineer’s coatings, and cathodic protection systems.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );

  pObj4 = docx.createNestedOrderedList({
    level: 1,
  });

  pObj4.addText(`Project Approach`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj4.addLineBreak();
  pObj4.addLineBreak();
  pObj4.addText(
    `We propose the following scope of services to maintain compliance with the SPCC Rules and Regulations:`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );

  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Conduct a site compliance evaluation and walk-down each oil storage container/tank.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Determine the type and construction standard of every oil storage container/tank.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Evaluate secondary containment, inspecting the imperviousness and ability to hold the largest container’s contents plus the rainwater from a 25-year 24-hour storm.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Evaluate the facility’s risk to navigable waters of the United States.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Establish tank integrity test intervals in accordance with applicable industry standards (STI-SP001 or API653.).`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Evaluate current spill kits and determine if they are adequate to contain the most likely spill outside containment.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Ensure overfill prevention measures are adequate and/or recommend tank upgrades.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `Draft a site map showing facility layout, spill control structures, and drainage patterns.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `We will write the SPCC Plan in accordance with most current state and federal oil pollution prevention regulations.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createNestedOrderedList({
    level: 2,
  });
  pObj4.addText(
    `The fieldwork will be supervised by Daniel Elliott and Pacific Engineering & Consulting personnel.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );

  pObj4 = docx.createNestedOrderedList({
    level: 1,
  });
  pObj4.addText(`Contractor Qualification`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj4.addLineBreak();
  pObj4.addText(`Jared Shuman – Pacific Engineering & Consulting`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addLineBreak();
  pObj4.addText(`Education:`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj4.addLineBreak();
  pObj4.addText(`MBA, California State University - Fresno - 2012`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`BS Mechanical Engineering, UCLA - 2008`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addLineBreak();
  pObj4.addText(`Expertise:`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj4.addLineBreak();
  pObj4.addText(
    `Certified Professional Engineer (PE) in the state of California – Certification #M36728`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4.addLineBreak();
  pObj4.addText(`Tank Inspector with certifications that include: `, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(
    `API-653 Aboveground Storage Tank Inspector – Certification #56100`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4.addLineBreak();
  pObj4.addText(
    `STI-001 Aboveground Storage Tank Inspector – Certification # 121286`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4.addLineBreak();

  pObj4.addText(
    `Mr. Shuman is a mechanical engineer experienced in storage tank structural analysis and in various hazardous material management and spill prevention planning and review processes.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );

  pObj4.addLineBreak();
  pObj4.addLineBreak();
  pObj4.addText(`Experience:`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj4.addLineBreak();
  pObj4.addText(
    `Consulting Engineer – 2010 to present: Lead engineer for Pacific Management Services / Pacific Engineering & Consulting specializing in environmental compliance inspections and planning. Has provided recommended updates to a range of environmental protection plans including Hazardous Materials Business Plans (HMBP), Storm Water Pollution Prevention Plans (SWPPP) and Spill Prevention Control and Countermeasure (SPCC) Plans. Evaluated petrochemical storage tanks in accordance with following applicable codes and standards:`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4.addLineBreak();
  pObj4.addText(`CAL EPA (CUPA, SWRCB, ARB, CalOSHA), `, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`American Petroleum Institute (API), `, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`Steel Tank Institute (STI),`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`American Society for Mechanical Engineers (ASME),`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`Underwriters Laboratory UL-142,`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`National Fire Prevention Association (NFPA) 30,`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`40 Code of Federal Regulation (CFR), and`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addText(`State and Federal regulation.`, {
    font_face: "Times New Roman",
    font_size: 12,
  });
  pObj4.addLineBreak();
  pObj4.addLineBreak();
  pObj4.addLineBreak();
  pObj4.addText(`Specific Related Projects:`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true,
  });
  pObj4 = docx.createListOfDots({
    level: 2,
  });
  pObj4.addText(
    `Inspector and Lead Engineer for Cleaning and Inspection of six ASTs at Air Force Plant 42. Conducted STI-SP001 inspections and certified calibration charts for all six tanks.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createListOfDots({
    level: 2,
  });
  pObj4.addText(
    `Lead inspector for STI-SP001 inspection digester tank for Las Gallinas Valley Sanitary District, which included an ultrasonic thickness test and engineering evaluation.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createListOfDots({
    level: 2,
  });
  pObj4.addText(
    `Inspector: Conducted Hazardous Waste Assessments of two used oil tanks for Cornerstone including ultrasonic thickness testing, pressure-decay test, seismic evaluation and a visual inspection of the tank appurtenances.`,
    {
      font_face: "Times New Roman",
      font_size: 12,
    }
  );
  pObj4 = docx.createListOfDots({
    level: 2,
  });
  pObj4.addText(`Inspected and evaluation of the Fresno Veterans Administration Hospital's petroleum storage tanks and developed procedures for spill response and emergency notification.`, {
    font_face: "Times New Roman",
    font_size: 12,
  });

  pObj4 = docx.createNestedOrderedList({
    level: 1,
  });
  pObj4.addText(`Project Cost/Limitations`, {
    font_face: "Times New Roman",
    font_size: 12,
    bold: true
  });


  // var table = [
  //   [
  //     {
  //       val: "No.",
  //       opts: {
  //         cellColWidth: 100,
  //         b: true,
  //         sz: "20",
  //         spacingBefore: 10,
  //         spacingAfter: 10,
  //         spacingLine: 20,
  //         spacingLineRule: "atLeast",
  //         shd: {
  //           fill: "7F7F7F",
  //           themeFill: "text1",
  //           themeFillTint: "80",
  //         },
  //         fontFamily: "Avenir Book",
  //       },
  //     },
  //     {
  //       val: "Title1",
  //       opts: {
  //         b: true,
  //         color: "A00000",
  //         align: "right",
  //         shd: {
  //           fill: "92CDDC",
  //           themeFill: "text1",
  //           themeFillTint: "80",
  //         },
  //       },
  //     },
  //   ],
  //   [1, "Col 1"],
  //   [2, "Col2"],
  //   [3, "Col 3"],
  //   [4, "Col 4"],
  //   [5, "Col 5"],
  //   [6, "Col 6"],
  // ];

  // var tableStyle = {
  //   tableColWidth: 100,
  //   tableSize: 20,
  //   tableAlign: "left",
  //   tableFontFamily: "Comic Sans MS",
  //   borders: true,
  //   borderSize: 2,
  //   columns: [{ width: 1 }, { width: 1 }, { width: 1 }], // Table logical columns
  // };

  // docx.createTable(table, tableStyle);
  let out = fs.createWriteStream("example.docx");

  out.on("error", function (err) {
    console.log(err);
  });

  docx.generate(out);
};
