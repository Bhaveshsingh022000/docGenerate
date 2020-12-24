const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Header,
  Footer,
  PageNumber,
  Media,
  UnderlineType,
  Numbering,
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeightRule,
} = require("docx");

exports.document2 = (req, res, next) => {
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "numberList",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 400,
                    hanging: 400,
                  },
                },
              },
            },
            {
              level: 1,
              format: "decimal",
              text: "%2.",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: 1500,
                    hanging: 400,
                  },
                },
              },
            },
          ],
        },
      ],
    },
  });
  const corpName = "Test Corp Name";
  const addressLine1 = "Test Address line 1";
  const addressLine2 = "Test Address line 2";
  const day = "22";
  const month = "12";
  const year = "2020";
  const totalEstimatedCost = "59.00";
  const priceBreakDownArray = [
    {
      "SPCC Plan": {
        quantity: "1",
        rate: "3,495.00",
        total: "0.00",
      },
    },
    {
      Lodging: {
        quantity: "0",
        rate: "150.00",
        total: "0.00",
      },
    },
    {
      Travel: {
        quantity: "0 hours",
        rate: "40.00",
        total: "0.00",
      },
    },
    {
      "Per Diem": {
        quantity: "1 day",
        rate: "55.00",
        total: "55.00",
      },
    },
    {
      Mileage: {
        quantity: "0 miles",
        rate: "0.55",
        total: "0.00",
      },
    },
  ];
  const pacificLogo = Media.addImage(
    doc,
    fs
      .readFileSync(path.resolve(__dirname, "../public/images/logo.png"))
      .toString("base64"),
    400,
    90
  );
  const pacificLogoHeader = Media.addImage(
    doc,
    fs
      .readFileSync(path.resolve(__dirname, "../public/images/logo.png"))
      .toString("base64"),
    350,
    80
  );

  doc.addSection({
    properties: {},
    children: [
      new Paragraph({
        spacing: {
          before: 900,
          after: 400,
        },
        children: [
          new TextRun({
            text: "Proposal:",
            bold: true,
            font: "Times New Roman",
            size: 24 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 400,
          after: 900,
        },
        children: [
          new TextRun({
            text:
              "Spill Control and Countermeasure Plan (SPCC Plan) Site Evaluation & SPCC Drafting",
            bold: true,
            font: "Times New Roman",
            size: 18 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 1800,
          after: 300,
        },
        children: [
          new TextRun({
            text: "Submitted To",
            italics: true,
            font: "Times New Roman",
            size: 14 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 400,
        },
        children: [
          new TextRun({
            text: `${corpName}`,
            bold: true,
            font: "Times New Roman",
            size: 18 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `${addressLine1}`,
            bold: true,
            font: "Times New Roman",
            size: 14 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `${addressLine2}`,
            bold: true,
            font: "Times New Roman",
            size: 14 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 500,
          after: 500,
        },
        children: [
          new TextRun({
            text: `${day} ${month} ${year}`,
            bold: true,
            font: "Times New Roman",
            size: 14 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 200,
          after: 200,
        },
        children: [pacificLogo],
        alignment: AlignmentType.CENTER,
      }),
    ],
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              pacificLogoHeader,
              new TextRun({
                text: "Environmental Engineering and Consulting",
                color: "#338ec4",
              }),
            ],
            alignment: AlignmentType.START,
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                children: [PageNumber.CURRENT],
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
      }),
    },
    properties: {},
    children: [
      new Paragraph({
        spacing: {
          before: 500,
          after: 200,
        },
        children: [
          new TextRun({
            text: "Pacific Engineering & Consulting",
            font: "Times New Roman",
            size: 14 * 2,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 200,
          after: 200,
        },
        children: [
          new TextRun({
            text: "Company Credentials and Certifications",
            font: "Times New Roman",
            size: 12 * 2,
            underline: UnderlineType.SINGLE,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "Name of Organization: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text: "Pacific Engineering & Consulting ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "Office Address: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text: "1788 N Helm Ave Suite 112 Fresno CA, 93727",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "Web Address: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text: "www.pacificmgt.com",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "Telephone Number: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text: "(559) 251-4060",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "Fax: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text: "(559) 251-4060",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "DUNS Number: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text: "196553770",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "Size of Company: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text:
              "Small under NAICS 541330 - Engineering Services, 541620 - Environmental Consulting Service,541690 - Other Scientific and Technical Consulting Services",
            font: "Times New Roman",
            size: 12 * 2,
            break: true,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          before: 800,
          after: 800,
        },
        children: [
          new TextRun({
            text: "Point of Contact: ",
            font: "Times New Roman",
            size: 12 * 2,
            bold: true,
          }),
          new TextRun({
            text:
              "Daniel Elliott Project Manager (559) 251-4060 ext. 105 daniel@pacificmgt.com ",
            font: "Times New Roman",
            size: 12 * 2,
            break: true,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
    ],
  });

  const tableRows = [];

  tableRows.push(
    new TableRow({
      children: [
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Task",
                  size: 12 * 2,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Quantity",
                  size: 12 * 2,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Rate",
                  size: 12 * 2,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Total",
                  size: 12 * 2,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
      ],
      height: {
        rule: HeightRule.EXACT,
        height: 300,
      },
    })
  );

  priceBreakDownArray.map((el) => {
    for (const [key, value] of Object.entries(el)) {
      const { quantity, rate, total } = value;
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: key,
                      size: 12 * 2,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: quantity,
                      size: 12 * 2,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: rate,
                      size: 12 * 2,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: total,
                      size: 12 * 2,
                    }),
                  ],
                }),
              ],
            }),
          ],
        })
      );
    }
  });

  tableRows.push(
    new TableRow({
      children: [
        new TableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Total Estimated Cost",
                  size: 12 * 2,
                }),
              ],
              alignment: AlignmentType.RIGHT
            }),
          ],
          columnSpan: 3
        }),
        new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: `${totalEstimatedCost}`,
                    size: 12 * 2,
                  }),
                ],
              }),
            ],
          }),
      ],
    })
  );

  const table = new Table({
    rows: tableRows,
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              pacificLogoHeader,
              new TextRun({
                text: "Environmental Engineering and Consulting",
                color: "#338ec4",
              }),
            ],
            alignment: AlignmentType.START,
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                children: [PageNumber.CURRENT],
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
      }),
    },
    children: [
      new Paragraph({
        spacing: {
          before: 400,
          after: 300,
        },
        children: [
          new TextRun({
            text: "Introduction: ",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
            break: true,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        spacing: {
          after: 500,
        },
        children: [
          new TextRun({
            text:
              "\tPacific Engineering & Consulting was founded in 1982 and is headquartered in Fresno, CA. We currently specialize in engineering and environmental services, specifically in performing certified inspection(s) and SPCC Plans. Pacific Engineering & Consulting has conducted Aboveground and Underground Storage Tank (AST/UST) inspections and testing for a variety of US Government agencies and Private Industry customers.  We have also done pressure testing, pressure vessel inspection, Spill Prevention Control and Countermeasure (SPCC) plans, Facility Response Plans (FRP). For government and commercial customers throughout the continental US. Pacific Engineering & Consulting is certified to conduct a variety of inspections including: storage tanks, piping, pressure vessels, OSHA’s Process Safety Management integrity inspections, Non-Destructive Examination, National Association of Corrosion Engineer’s coatings, and cathodic protection systems.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "Project Approach: ",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "We propose the following scope of services to maintain compliance with the SPCC Rules and Regulations: ",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Conduct a site compliance evaluation and walk-down each oil storage container/tank.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Determine the type and construction standard of every oil storage container/tank.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Evaluate secondary containment, inspecting the imperviousness and ability to hold the largest container’s contents plus the rainwater from a 25-year 24-hour storm.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Evaluate the facility’s risk to navigable waters of the United States.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Establish tank integrity test intervals in accordance with applicable industry standards (STI-SP001 or API653.)",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Evaluate current spill kits and determine if they are adequate to contain the most likely spill outside containment.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Ensure overfill prevention measures are adequate and/or recommend tank upgrades.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Draft a site map showing facility layout, spill control structures, and drainage patterns.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "We will write the SPCC Plan in accordance with most current state and federal oil pollution prevention regulations.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "The fieldwork will be supervised by Daniel Elliott and Pacific Engineering & Consulting personnel.",
            bold: false,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        pageBreakBefore: true,
        spacing: {
          before: 400,
        },
        children: [
          new TextRun({
            text: "Contractor Qualification",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "Jared Shuman – Pacific Engineering & Consulting",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Education:",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "MBA, California State University - Fresno - 2012",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "BS Mechanical Engineering, UCLA - 2008 ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Expertise: ",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Certified Professional Engineer (PE) in the state of California – Certification #M36728",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Tank Inspector with certifications that include: ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "\tAPI-653 Aboveground Storage Tank Inspector – Certification #56100",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 800,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "\tSTI-001 Aboveground Storage Tank Inspector – Certification # 121286",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 800,
        },
      }),
      new Paragraph({
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text:
              "Mr. Shuman is a mechanical engineer experienced in storage tank structural analysis and in various hazardous material management and spill prevention planning and review processes.",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Experience:",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Consulting Engineer – 2010 to present: Lead engineer for Pacific Management Services / Pacific Engineering & Consulting specializing in environmental compliance inspections and planning. Has provided recommended updates to a range of environmental protection plans including Hazardous Materials Business Plans (HMBP), Storm Water Pollution Prevention Plans (SWPPP) and Spill Prevention Control and Countermeasure (SPCC) Plans. Evaluated petrochemical storage tanks in accordance with following applicable codes and standards:",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "\tCAL EPA (CUPA, SWRCB, ARB, CalOSHA), ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "\tAmerican Petroleum Institute (API), ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "\tSteel Tank Institute (STI),",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "\tAmerican Society for Mechanical Engineers (ASME), ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "\tUnderwriters Laboratory UL-142",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "\tNational Fire Prevention Association (NFPA) 30, ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "\t40 Code of Federal Regulation (CFR), and",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "\tState and Federal regulation.",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Specific Related Projects:",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          left: 600,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Inspector and Lead Engineer for Cleaning and Inspection of six ASTs at Air Force Plant 42. Conducted STI-SP001 inspections and certified calibration charts for all six tanks",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Lead inspector for STI-SP001 inspection digester tank for Las Gallinas Valley Sanitary District, which included an ultrasonic thickness test and engineering evaluation",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Inspector: Conducted Hazardous Waste Assessments of two used oil tanks for Cornerstone including ultrasonic thickness testing, pressure-decay test, seismic evaluation and a visual inspection of the tank appurtenances.",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "Inspected and evaluation of the Fresno Veterans Administration Hospital's petroleum storage tanks and developed procedures for spill response and emergency notification. ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        spacing: {
          before: 300,
          after: 300,
        },
        children: [
          new TextRun({
            text: "Project Cost/Limitations",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              "The estimated cost to provide the above-mentioned engineering services is $0.00.",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: "A breakdown of the estimated fees is provided below:",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      table,
      new Paragraph({
        spacing: { before: 200, after: 200 },
        children: [
          new TextRun({
            text:
              "If there are any questions or concerns regarding the proposed services and associated fees, please do not hesitate to contact Pacific Engineering & Consulting at your earliest convenience. Pacific Engineering & Consulting strives to satisfy our client’s needs and meet their expectations. We will make every effort to accommodate requested changes in our understanding of the project, assumptions, scope, or services, as appropriate. ",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [
          new TextRun({
            text:
              "Proposed costs are good for 60 days from the date of issue noted above.",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        children: [
          new TextRun({
            text:
              'Extra work, if required, will be completed on a separately negotiated lump sum basis or on a "time and materials" basis according to Pacific Engineering & Consulting’s fee schedule. No extra work will be performed without written authorization from the client.',
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
    ],
  });

  doc.addSection({
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              pacificLogoHeader,
              new TextRun({
                text: "Environmental Engineering and Consulting",
                color: "#338ec4",
              }),
            ],
            alignment: AlignmentType.START,
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({
                children: [PageNumber.CURRENT],
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
      }),
    },
    children: [
      new Paragraph({
        spacing: { before: 300, after: 200 },
        children: [
          new TextRun({
            text: "Conclusion: ",
            bold: true,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        spacing: {
          after: 200,
        },
        children: [
          new TextRun({
            text: `Pacific Engineering & Consulting will administer this project in accordance with all applicable ${corpName} requirements, industry standards, and engineering best practices. Our staff of Professional Engineers and certified personnel are excited at the opportunity to assist Corp Name environmental compliance needs.`,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        spacing: {
          after: 600,
        },
        children: [
          new TextRun({
            text: "Best Regards",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Daniel Elliott",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "STI Inspector, #AC44220",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "API 653 Inspector, #70788",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "API 570 Inspector, #82919",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "QISP #00969",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        spacing: { after: 900 },
        children: [
          new TextRun({
            text: "NDT Level II",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: "Please Sign and Return to Pacific Engineering & Consulting.",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Accepted by:",
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        spacing: { after: 450 },
        children: [
          new TextRun({
            text: `${corpName} / Authorized Representative`,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        spacing: { after: 450 },
        children: [
          new TextRun({
            text: `Signature`,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Title \t\t Date`,
            font: "Times New Roman",
            size: 12 * 2,
          }),
        ],
      }),
    ],
  });

  Packer.toBuffer(doc)
    .then((buffer) => {
      fs.writeFileSync("My Document.docx", buffer);
    })
    .then(() => {
      res.send("Done");
      res.end();
    });
};
