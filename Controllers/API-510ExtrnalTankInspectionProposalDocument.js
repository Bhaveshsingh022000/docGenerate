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
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeightRule,
  VerticalAlign,
  BorderStyle,
  RelativeVerticalPosition,
  TableLayoutType,
  PageBorders,
  PageBorderDisplay,
} = require("docx");
const { type } = require("os");

exports.proposalDocument = (req, res, next) => {
  const headerFontColor = "#338ec4";
  const headerBorderBottomColor = "#2f793b";
  const contentFontSize = 12 * 2;
  const contentHeadingFontSize = 14 * 2;
  const contentFont = "Times New Roman";
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
                run: {
                  bold: true,
                  size: contentFontSize,
                },
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
  const date = new Date();
  const corpName = "Ed Staub & Sons Petroleum, Inc.";
  const addressLine1 = "Test Address line 1";
  const addressLine2 = "Klamath Falls, OR 97601";
  const day = date.getDate();
  const month = date.getMonth();
  const year = date.getFullYear();
  const totalEstimatedCost = "59.00";
  const priceBreakDownArray = [
    {
      "API-510": {
        quantity: "1",
        rate: "3,000.00",
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

  const additionalCostsArray = [
    {
      "Articulating Boom Lift Rental (If necessary)": {
        quantity: "0",
        rate: "800.55",
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
      .readFileSync(path.resolve(__dirname, "../public/images/headerLogo.png"))
      .toString("base64"),
    330,
    60
  );

  const pacificCredentialImage = Media.addImage(
    doc,
    fs.readFileSync(
      path
        .resolve(__dirname, "../public/images/api150ExternalTankInspection.PNG")
        .toString("base64")
    ),
    210,
    650
  );

  const pageBorderProperties = {
    color: headerFontColor,
    size: 200,
    style: BorderStyle.THICK_THIN_SMALL_GAP,
    space: 200,
  };

  doc.addSection({
    margins: {
      left: 700,
      right: 700,
    },
    properties: {
      pageBorders: PageBorderDisplay.FIRST_PAGE,
      pageBorderTop: { ...pageBorderProperties },
      pageBorderBottom: { ...pageBorderProperties },
      pageBorderLeft: { ...pageBorderProperties },
      pageBorderRight: { ...pageBorderProperties },
    },
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
            font: contentFont,
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
            text: "API 510 In-Service Tank Inspection",
            bold: true,
            font: contentFont,
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
            font: contentFont,
            size: contentHeadingFontSize,
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
            font: contentFont,
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
            font: contentFont,
            size: contentHeadingFontSize,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `${addressLine2}`,
            bold: true,
            font: contentFont,
            size: contentHeadingFontSize,
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
            font: contentFont,
            size: contentHeadingFontSize,
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

  const companyCredentialsTableBorders = {
    left: { color: "white" },
    right: { color: "white" },
    bottom: { color: "white" },
    top: { color: "white" },
  };

  const companyCredentialsTable = new Table({
    float: {
      absoluteHorizontalPosition: 300,
    },
    width: {
      size: 70,
      type: WidthType.PERCENTAGE,
    },
    rows: [
      new TableRow({
        height: {
          height: 900,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            margins: {
              bottom: 100,
            },
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Name of Organization: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
            width: {
              size: 900 * 4,
              type: WidthType.DXA,
            },
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Pacific Engineering & Consulting ",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
            width: {
              size: 600 * 4,
              type: WidthType.DXA,
            },
          }),
        ],
      }),
      new TableRow({
        height: {
          height: 900,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            margins: {
              bottom: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Office Address: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "1788 N Helm Ave Suite 112 Fresno CA, 93727",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        height: {
          height: 900,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            margins: {
              bottom: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Web Address: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "www.pacificmgt.com",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        height: {
          height: 900,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            margins: {
              bottom: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Telephone Number: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "(559) 251-4060",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        height: {
          height: 900,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            margins: {
              bottom: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Fax: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "(559) 251-4060",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        height: {
          height: 900,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            margins: {
              bottom: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "DUNS Number: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "196553770",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        height: {
          height: 1500,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            margins: {
              bottom: 600,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Size of Company: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Small under NAICS",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "541330 - Engineering Services",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "541620 - Environmental Consulting Service",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text:
                      "541690 - Other Scientific and Technical Consulting Services",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        height: {
          height: 1200,
          rule: HeightRule.EXACT,
        },
        children: [
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            margins: {
              bottom: 100,
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Point of Contact: ",
                    font: contentFont,
                    size: contentFontSize,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: { ...companyCredentialsTableBorders },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Daniel Elliott",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Project Manager",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "(559) 251-4060 ext. 105",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "daniel@pacificmgt.com",
                    font: contentFont,
                    size: contentFontSize,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  doc.addSection({
    margins: {
      left: 700,
      right: 700,
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: {
                      size: 1500 * 4,
                      type: WidthType.DXA,
                    },
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [pacificLogoHeader],
                      }),
                    ],
                  }),
                  new TableCell({
                    margins: {
                      left: 600,
                    },
                    verticalAlign: VerticalAlign.BOTTOM,
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Environmental Engineering and Consulting",
                            color: headerFontColor,
                            break: true,
                            size: 11 * 2,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
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
            font: contentFont,
            size: contentHeadingFontSize,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        spacing: {
          before: 200,
          after: 700,
        },
        children: [
          new TextRun({
            text: "Company Credentials and Certifications",
            font: contentFont,
            size: contentFontSize,
            underline: UnderlineType.SINGLE,
          }),
        ],
        alignment: AlignmentType.CENTER,
      }),
      companyCredentialsTable,
      new Paragraph({
        children: [pacificCredentialImage],
      }),
    ],
  });

  const costTableRows = [];
  const priceBreakdownTableCellMargin = {
    left: 100,
    right: 100,
  };

  costTableRows.push(
    new TableRow({
      children: [
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          width: {
            size: 600 * 4,
            type: WidthType.PERCENTAGE,
          },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Task",
                  size: contentFontSize,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Quantity",
                  size: contentFontSize,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Rate",
                  size: contentFontSize,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Total",
                  size: contentFontSize,
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

  priceBreakDownArray.map((element) => {
    for (const [key, value] of Object.entries(element)) {
      const { quantity, rate, total } = value;
      costTableRows.push(
        new TableRow({
          children: [
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: key,
                      size: contentFontSize,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: quantity,
                      size: contentFontSize,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: `$${rate}`,
                      size: contentFontSize,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: `$${total}`,
                      size: contentFontSize,
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

  costTableRows.push(
    new TableRow({
      children: [
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Total Estimated Cost",
                  size: contentFontSize,
                }),
              ],
              alignment: AlignmentType.RIGHT,
            }),
          ],
          columnSpan: 3,
        }),
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: `$${totalEstimatedCost}`,
                  size: contentFontSize,
                }),
              ],
            }),
          ],
        }),
      ],
    })
  );

  const costTable = new Table({
    float: {
      absoluteHorizontalPosition: 500,
      verticalAnchor: 100,
    },
    rows: costTableRows,
    width: {
      size: 90,
      type: WidthType.PERCENTAGE,
    },
    margins: {
      bottom: 1500,
    },
  });

  const additionalCostsTableRows = [];

  additionalCostsTableRows.push(
    new TableRow({
      children: [
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          width: {
            size: 600 * 4,
            type: WidthType.PERCENTAGE,
          },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Additional Costs",
                  size: contentFontSize,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Quantity",
                  size: contentFontSize,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Rate",
                  size: contentFontSize,
                  bold: true,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          margins: { ...priceBreakdownTableCellMargin },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Total",
                  size: contentFontSize,
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

  additionalCostsArray.map((element) => {
    for (const [key, value] of Object.entries(element)) {
      const { quantity, rate, total } = value;
      additionalCostsTableRows.push(
        new TableRow({
          children: [
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: key,
                      size: contentFontSize,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: quantity,
                      size: contentFontSize,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: `$${rate}`,
                      size: contentFontSize,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              margins: { ...priceBreakdownTableCellMargin },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: `$${total}`,
                      size: contentFontSize,
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

  const additionalCostsTable = new Table({
    float: {
      absoluteHorizontalPosition: 500,
      absoluteVerticalPosition: 4200,
    },

    rows: additionalCostsTableRows,
    width: {
      size: 90,
      type: WidthType.PERCENTAGE,
    },
  });

  const contentIndent = {
    left: 1200,
    right: 300,
  };

  doc.addSection({
    margins: {
      left: 700,
      right: 700,
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: {
                      size: 1500 * 4,
                      type: WidthType.DXA,
                    },
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [pacificLogoHeader],
                      }),
                    ],
                  }),
                  new TableCell({
                    margins: {
                      left: 600,
                    },
                    verticalAlign: VerticalAlign.BOTTOM,
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Environmental Engineering and Consulting",
                            color: headerFontColor,
                            break: true,
                            size: 11 * 2,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
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
        indent: {
          ...contentIndent,
        },
        spacing: {
          before: 400,
          after: 300,
        },
        children: [
          new TextRun({
            text: "Introduction ",
            bold: true,
            font: contentFont,
            size: contentFontSize,
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
        indent: {
          ...contentIndent,
        },
        spacing: {
          after: 500,
        },
        children: [
          new TextRun({
            text:
              "\tPacific Engineering & Consulting was founded in 1982 and is headquartered in Fresno, CA. We currently specialize in engineering and environmental services, specifically in performing certified inspection(s) and SPCC Plans. Pacific Engineering & Consulting has conducted Aboveground and Underground Storage Tank (AST/UST) inspections and testing for a variety of US Government agencies and Private Industry customers.  We have also done pressure testing, pressure vessel inspection, Spill Prevention Control and Countermeasure (SPCC) plans, Facility Response Plans (FRP). For government and commercial customers throughout the continental US. Pacific Engineering & Consulting is certified to conduct a variety of inspections including: storage tanks, piping, pressure vessels, OSHA’s Process Safety Management integrity inspections, Non-Destructive Examination, National Association of Corrosion Engineer’s coatings, and cathodic protection systems.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
        },
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "Project Approach: ",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
        },
        children: [
          new TextRun({
            text:
              "We propose the following scope of services for an in-service inspection of the pressure vessel per API 510:",

            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1600,
        },
        children: [
          new TextRun({
            text:
              "Review U-1 sheet, installation documents and previous inspection reports, if available for each vessel.",

            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1600,
        },
        children: [
          new TextRun({
            text: "Visually inspect vessel exterior per API-510. ",

            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1600,
        },
        children: [
          new TextRun({
            text:
              "Ultrasonic thickness evaluation of the tank jacket exterior. ",

            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1600,
        },
        children: [
          new TextRun({
            text:
              "Establish vessel test intervals in accordance with API 510. ",

            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1600,
        },
        children: [
          new TextRun({
            text:
              "Prepare a written report meeting the requirements of API-510 including our findings and recommendations and an estimate of the vessel remaining life. ",

            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1600,
        },
        children: [
          new TextRun({
            text:
              "The fieldwork will be supervised by Daniel Elliott and Pacific Engineering & Consulting personnel. We anticipate needed no more than one (1) day on-site",

            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1600,
        },
        pageBreakBefore: true,
        spacing: {
          before: 400,
          after: 200,
        },
        children: [
          new TextRun({
            text: "Contractor Qualification",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "Jared Shuman – Pacific Engineering & Consulting",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        children: [
          new TextRun({
            text: "Education:",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        children: [
          new TextRun({
            text: "MBA, California State University - Fresno - 2012",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "BS Mechanical Engineering, UCLA - 2008 ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        children: [
          new TextRun({
            text: "Expertise: ",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        children: [
          new TextRun({
            text:
              "Certified Professional Engineer (PE) in the state of California – Certification #M36728",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        children: [
          new TextRun({
            text: "Tank Inspector with certifications that include: ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
        },
        children: [
          new TextRun({
            text:
              "API-653 Aboveground Storage Tank Inspector – Certification #56100",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
        },
        children: [
          new TextRun({
            text:
              "STI-001 Aboveground Storage Tank Inspector – Certification # 121286",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text:
              "Mr. Shuman is a mechanical engineer experienced in storage tank structural analysis and in various hazardous material management and spill prevention planning and review processes.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        children: [
          new TextRun({
            text: "Experience:",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 500,
        },
        children: [
          new TextRun({
            text:
              "Consulting Engineer – 2010 to present: Lead engineer for Pacific Management Services / Pacific Engineering & Consulting specializing in environmental compliance inspections and planning. Has provided recommended updates to a range of environmental protection plans including Hazardous Materials Business Plans (HMBP), Storm Water Pollution Prevention Plans (SWPPP) and Spill Prevention Control and Countermeasure (SPCC) Plans. Evaluated petrochemical storage tanks in accordance with following applicable codes and standards:",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 900,
        },
        children: [
          new TextRun({
            text: "CAL EPA (CUPA, SWRCB, ARB, CalOSHA), ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 900,
        },
        children: [
          new TextRun({
            text: "American Petroleum Institute (API), ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Steel Tank Institute (STI),",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          ...contentIndent,
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "American Society for Mechanical Engineers (ASME), ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          ...contentIndent,
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Underwriters Laboratory UL-142",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          ...contentIndent,
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "National Fire Prevention Association (NFPA) 30, ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          ...contentIndent,
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "40 Code of Federal Regulation (CFR), and",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          ...contentIndent,
          left: 900,
        },
      }),
      new Paragraph({
        spacing: {
          after: 300,
        },
        children: [
          new TextRun({
            text: "State and Federal regulation.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          ...contentIndent,
          left: 900,
        },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: "Specific Related Projects:",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        indent: {
          ...contentIndent,
          left: 500,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1300,
        },
        children: [
          new TextRun({
            text:
              "Inspector and Lead Engineer for Cleaning and Inspection of six ASTs at Air Force Plant 42. Conducted STI-SP001 inspections and certified calibration charts for all six tanks",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1300,
        },
        children: [
          new TextRun({
            text:
              "Lead inspector for STI-SP001 inspection digester tank for Las Gallinas Valley Sanitary District, which included an ultrasonic thickness test and engineering evaluation",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1300,
        },
        children: [
          new TextRun({
            text:
              "Inspector: Conducted Hazardous Waste Assessments of two used oil tanks for Cornerstone including ultrasonic thickness testing, pressure-decay test, seismic evaluation and a visual inspection of the tank appurtenances.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1300,
        },
        spacing: {
          after: 9000,
        },
        children: [
          new TextRun({
            text:
              "Inspected and evaluation of the Fresno Veterans Administration Hospital's petroleum storage tanks and developed procedures for spill response and emergency notification. ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 1,
        },
      }),
    ],
  });

  doc.addSection({
    margins: {
      left: 700,
      right: 700,
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: {
                      size: 1500 * 4,
                      type: WidthType.DXA,
                    },
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [pacificLogoHeader],
                      }),
                    ],
                  }),
                  new TableCell({
                    margins: {
                      left: 600,
                    },
                    verticalAlign: VerticalAlign.BOTTOM,
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Environmental Engineering and Consulting",
                            color: headerFontColor,
                            break: true,
                            size: 11 * 2,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
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
        indent: {
          ...contentIndent,
        },
        spacing: { before: 300, after: 200 },
        children: [
          new TextRun({
            text: "Project Cost/Limitations",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: `The estimated cost to provide the above-mentioned engineering services is $${totalEstimatedCost}.`,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { after: 300 },
        children: [
          new TextRun({
            text: "A breakdown of the estimated fees is provided below:",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      costTable,
      additionalCostsTable,
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { before: 100 },
        children: [
          new TextRun({
            text: "Assumptions",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1200,
        },
        children: [
          new TextRun({
            text:
              "Confined space entry, if applicable, are gas free, safe for entry, and accessible using natural ventilation only and accessible using fixed ladders.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        bullet: {
          level: 0,
        },
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 1200,
        },
        spacing: { after: 300 },
        children: [
          new TextRun({
            text:
              "Any previous inspection reports, engineering drawings, reports or calculations will be provided upon request.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        bullet: {
          level: 0,
        },
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { before: 100, after: 300 },
        children: [
          new TextRun({
            text:
              "If there are any questions or concerns regarding the proposed services and associated fees, please do not hesitate to contact Pacific Engineering & Consulting at your earliest convenience. Pacific Engineering & Consulting strives to satisfy our client’s needs and meet their expectations. We will make every effort to accommodate requested changes in our understanding of the project, assumptions, scope, or services, as appropriate. ",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { after: 300 },
        children: [
          new TextRun({
            text:
              "Proposed costs are good for 60 days from the date of issue noted above.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text:
              'Extra work, if required, will be completed on a separately negotiated lump sum basis or on a "time and materials" basis according to Pacific Engineering & Consulting’s fee schedule. No extra work will be performed without written authorization from the client.',
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        alignment: AlignmentType.LEFT,
      }),
    ],
  });

  doc.addSection({
    margins: {
      left: 700,
      right: 700,
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: {
                      size: 1500 * 4,
                      type: WidthType.DXA,
                    },
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [pacificLogoHeader],
                      }),
                    ],
                  }),
                  new TableCell({
                    margins: {
                      left: 600,
                    },
                    verticalAlign: VerticalAlign.BOTTOM,
                    borders: {
                      ...companyCredentialsTableBorders,
                      bottom: {
                        color: headerBorderBottomColor,
                        size: 20,
                        style: BorderStyle.THICK,
                      },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Environmental Engineering and Consulting",
                            color: headerFontColor,
                            break: true,
                            size: 11 * 2,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
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
        indent: {
          ...contentIndent,
        },
        spacing: { before: 300, after: 200 },
        children: [
          new TextRun({
            text: "Conclusion ",
            bold: true,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
        numbering: {
          reference: "numberList",
          level: 0,
        },
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: {
          after: 200,
        },
        children: [
          new TextRun({
            text: `Pacific Engineering & Consulting will administer this project in accordance with all applicable ${corpName} requirements, industry standards, and engineering best practices. Our staff of Professional Engineers and certified personnel are excited at the opportunity to assist ${corpName} environmental compliance needs.`,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: {
          after: 600,
        },
        children: [
          new TextRun({
            text: "Best Regards",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: "Daniel Elliott",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: "STI Inspector, #AC44220",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: "API 653 Inspector, #70788",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: "API 570 Inspector, #82919",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: "QISP #00969",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { after: 2000 },
        children: [
          new TextRun({
            text: "NDT Level II",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { before: 400, after: 300 },
        children: [
          new TextRun({
            text: "Please Sign and Return to Pacific Engineering & Consulting.",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: "Accepted by:",
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { after: 450 },
        children: [
          new TextRun({
            text: `${corpName} / Authorized Representative`,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: `____________________________________`,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        spacing: { after: 450 },
        children: [
          new TextRun({
            text: `Signature`,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: `____________________________________`,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
      new Paragraph({
        indent: {
          ...contentIndent,
          left: 400,
        },
        children: [
          new TextRun({
            text: `Title \t\t\t\t\t Date`,
            font: contentFont,
            size: contentFontSize,
          }),
        ],
      }),
    ],
  });

  Packer.toBuffer(doc)
    .then((buffer) => {
      fs.writeFileSync(
        "API-510 Extrnal Tank Inspection ProposalDocument.docx",
        buffer
      );
    })
    .then(() => {
      res.send("Done");
      res.end();
    });
};
