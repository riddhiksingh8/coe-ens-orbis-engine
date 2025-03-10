import * as fs from 'fs';
import {
  AlignmentType,
  BorderStyle,
  ImageRun,
  Paragraph,
  patchDocument,
  PatchType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import topdf from 'docx2pdf-converter';

import { BlobServiceClient } from '@azure/storage-blob';
import { getOrdinalSuffix, getRiskColor } from './helpers.js';

const TEMPLATE_PATH = 'src/template/aramco_template.docx';

// Azure Storage Credentials
const { BLOB_STORAGE__CONNECTION_STRING } = process.env;

// Initialize Azure Blob Service Client
const blobServiceClient = BlobServiceClient.fromConnectionString(
  BLOB_STORAGE__CONNECTION_STRING,
);

// Function to upload a file to Azure Blob Storage
const uploadToAzure = async (filePath, fileName, session_id) => {
  try {
    const containerClient = blobServiceClient.getContainerClient(session_id);
    const blobClient = containerClient.getBlockBlobClient(fileName);
    const fileBuffer = fs.readFileSync(filePath);

    await blobClient.uploadData(fileBuffer);
    return blobClient.url;
  } catch (error) {
    console.error(`Error uploading ${fileName} to Azure:`, error);
  }
};
const date = new Date();
const day = date.getDate();
const month = date.toLocaleString('en-US', { month: 'long' });
const year = date.getFullYear();

const ordinalSuffix = getOrdinalSuffix(day);

const borders = {
  top: {
    style: BorderStyle.SINGLE,
    size: 1,
    color: 'D3D3D3',
  }, // Light gray
  bottom: {
    style: BorderStyle.SINGLE,
    size: 1,
    color: 'D3D3D3',
  },
  left: {
    style: BorderStyle.SINGLE,
    size: 1,
    color: 'D3D3D3',
  },
  right: {
    style: BorderStyle.SINGLE,
    size: 1,
    color: 'D3D3D3',
  },
};

const createTextRun = (options) => ({
  type: PatchType.PARAGRAPH,
  children: [new TextRun(options)],
});

const createCell = (
  text,
  {
    background = 'F2F2F2',
    alignment = 'center',
    bold = true,
    columnSpan = 1,
  } = {},
) => {
  return new TableCell({
    verticalAlign: 'center',
    children: [
      new Paragraph({
        alignment,
        children: [new TextRun({ text, bold, size: 20 })],
      }),
    ],
    shading: { fill: background },
    columnSpan: columnSpan || 1,
  });
};

const highlightRating = (rating) => ({
  type: PatchType.DOCUMENT,
  children: [
    new Paragraph({
      alignment: 'center',
      shading: { fill: getRiskColor(rating).background },
      spacing: {
        line: 180,
      },
      children: [
        new TextRun({
          break: 1,
        }),
        new TextRun({
          text: `${rating}`,
          color: getRiskColor(rating).color,
          bold: true,
        }),
        new TextRun({
          break: 1,
        }),
      ],
    }),
  ],
});

const createNoHitsTable = (text = '') => {
  return [
    new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },

      rows: [
        new TableRow({
          height: { rule: 'atLeast', value: 500 },
          children: [
            new TableCell({
              verticalAlign: 'center',
              shading: {
                fill: 'F2F2F2',
              },
              children: [
                new Paragraph({
                  alignment: 'center',
                  // spacing: {
                  //   after: 100,
                  //   before: 100,
                  // },
                  children: [
                    new TextRun({
                      text: text
                        ? `${text} - NO TRUE HITS IDENTIFIED`
                        : 'NO TRUE HITS IDENTIFIED',
                      bold: true,
                      size: 20,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
};

const createFindingsInnerTable = (findings) => {
  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },

      rows: [
        // Header row
        new TableRow({
          height: { rule: 'atLeast', value: 500 },

          children: [
            createCell('Name & Relation', {
              background: 'f2f2f2',
              alignment: 'left',
            }),
            createCell(findings.title, {
              background: 'ffffff',
              alignment: 'center',
              bold: false,
            }),
            createCell('Rating'),
            createCell(findings.rating, {
              background: 'ffffff',
              alignment: 'center',
              bold: false,
            }),
          ],
        }),
        // Findings row (merged across 4 columns)
        new TableRow({
          height: { rule: 'atLeast', value: 500 },
          children: [
            createCell('Findings', {
              columnSpan: 4,
              alignment: 'left',
              background: 'f2f2f2',
            }),
          ],
        }),
        new TableRow({
          height: { rule: 'atLeast', value: 500 },
          children: [
            new TableCell({
              shading: { fill: 'ffffff' },
              columnSpan: 4,

              children: [
                new Paragraph({}),
                new Table({
                  width: { size: 100, type: WidthType.PERCENTAGE },
                  rows: [
                    new TableRow({
                      height: { rule: 'atLeast', value: 500 },
                      children: [
                        createCell(findings.inner_title),
                        createCell('Rating'),
                        createCell('Notes'),
                      ],
                    }),

                    ...findings.data.map((item) => {
                      return new TableRow({
                        height: { rule: 'atLeast', value: 500 },
                        children: [
                          createCell(item.kpi_definition, {
                            background: 'ffffff',
                            bold: false,
                          }),
                          createCell(item.kpi_rating, {
                            background: 'ffffff',
                            bold: false,
                          }),
                          createCell(item.kpi_details, {
                            background: 'ffffff',
                            bold: false,
                          }),
                        ],
                      });
                    }),
                  ],
                }),
                new Paragraph({}),
              ],
            }),
          ],
        }),
      ],
    }),
    new Paragraph({}),
  ];
};

const createFindingsTable = (findings) => {
  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      alignment: 'center',

      rows: [
        // Header row
        new TableRow({
          height: { rule: 'atLeast', value: 500 },
          children: [
            createCell('Name & Relation', {
              background: 'f2f2f2',
              alignment: 'left',
              bold: true,
            }),
            createCell(findings.kpi_definition, {
              background: 'ffffff',
              alignment: 'center',
              bold: false,
            }),
            createCell('Rating'),
            createCell(findings.kpi_rating, {
              background: 'ffffff',
              alignment: 'center',
              bold: false,
            }),
          ],
        }),
        // Findings row (merged across 4 columns)
        new TableRow({
          height: { rule: 'atLeast', value: 500 },
          // borders: tableBorders,
          children: [
            createCell('Findings', {
              columnSpan: 4,
              alignment: 'left',
              background: 'f2f2f2',
              bold: true,
            }),
          ],
        }),
        new TableRow({
          height: { rule: 'atLeast', value: 500 },
          children: [
            new TableCell({
              children: findings.kpi_details
                .trim()
                .split(/\n+/)
                .map(
                  (text) =>
                    new Paragraph({
                      spacing: {
                        after: 50,
                        before: 50,
                      },
                      children: [new TextRun({ text, bold: false })],
                    }),
                ),

              shading: { fill: 'ffffff' },
              columnSpan: 4,
            }),
          ],
        }),
      ],
    }),
    new Paragraph({}),
  ];
};

export const generateReport = async (payload) => {
  try {
    const data = {
      ...payload,
      riskData: [
        { area: 'Sanctions', rating: payload.sanctions_rating },
        {
          area: 'Anti-Bribery and Anti-Corruption',
          rating: payload.anti_rating,
        },
        {
          area: 'Government Ownership and Political Affiliations',
          rating: payload.gov_rating,
        },
        { area: 'Financial Indicators', rating: payload.financial_rating },
        { area: 'Other Adverse Media', rating: payload.adv_rating },
        { area: 'Cyber Security', rating: payload.cyber_rating },
        {
          area: 'ESG',
          rating: payload.esg_rating,
        },
        {
          area: 'Regulatory & Legal',
          rating: payload.regulatory_and_legal_rating,
        },
      ],
      riskAreas: {
        sanctions: payload.sanctions_summary,
        antiBriberyAndAntiCorruption: payload.anti_summary,
        governmentOwnershipAndPoliticalAffiliations: payload.gov_summary,
        financialIndicators: payload.financial_summary,
        otherAdverseMedia: payload.adv_summary,
        cyberSecurity: payload.cyber_summary,
        esg: payload.esg_summary,
        regulatoryAndLegal: payload.ral_summary,
      },
      cyberSecurity_findings: {
        title: `${payload.name} (Self)`,
        rating: payload.cyber_rating,
        inner_title: 'Cyber Security Indicators',
        data: payload.cyb_findings ? payload.cyb_data : [],
      },
      esg_findings: {
        title: `${payload.name} (Self)`,
        rating: payload.esg_rating,
        inner_title: 'ESG Indicators',
        data: payload.esg_findings ? payload.esg_data : [],
      },
    };

    console.log('Data:', data);
    const doc = await patchDocument({
      outputType: 'nodebuffer',
      data: fs.readFileSync(TEMPLATE_PATH),
      patches: {
        title: {
          type: PatchType.PARAGRAPH,
          children: [
            new TextRun(data.name),
            new ImageRun({
              type: 'png',
              data: fs.readFileSync('src/images/titleBackground.png'),
              transformation: { width: 500, height: 400 },
              floating: {
                behindDocument: true,
                horizontalPosition: {
                  relative: 'column',
                  align: 'left',
                },
                verticalPosition: {
                  offset: 705789,
                },
              },
            }),
          ],
        },
        created_date: {
          type: PatchType.PARAGRAPH,
          children: [
            new TextRun({ text: `${day}`, bold: true }), // Day (bold)
            new TextRun({ text: ordinalSuffix, superScript: true }), // Ordinal suffix (superscript)
            new TextRun({ text: ` ${month} ${year}` }), // Month and year
          ],
        },

        // Company Profile
        company_name: createTextRun({
          text: data.name,
        }),
        company_location: createTextRun({
          text: data.location,
        }),
        company_address: createTextRun({
          text: data.address,
        }),
        company_website: createTextRun({
          text: data.website,
        }),
        company_active_status: createTextRun({
          text: data.active_status,
        }),
        company_operation_type: createTextRun({
          text: data.operation_type,
        }),
        company_legal_status: createTextRun({
          text: data.legal_status,
        }),
        company_national_identifier: createTextRun({
          text: data.national_id,
        }),
        company_alias: createTextRun({
          text: data.alias,
        }),
        company_incorporation_date: createTextRun({
          text: data.incorporation_date,
        }),

        company_subsidiaries: createTextRun({
          text: data.subsidiaries,
        }),
        company_corporate_group: createTextRun({
          text: data.corporate_group,
        }),

        shareholders: {
          type: PatchType.DOCUMENT,
          children: data.shareholders
            .split('\n')
            .map((shareholder) => new Paragraph(shareholder)),
        },
        key_executives: {
          type: PatchType.DOCUMENT,
          children: data.key_exec
            .split('\n')
            .map((executive) => new Paragraph(executive)),
        },
        company_revenue: createTextRun({
          text: data.revenue,
        }),
        company_employee: createTextRun({
          text: data.employee_count,
        }),

        overall_rating: {
          type: PatchType.DOCUMENT,
          children: [
            new Table({
              columnWidths: [8000, 4000],
              width: {
                size: 70,
                type: WidthType.PERCENTAGE,
              },
              alignment: 'center',

              rows: [
                new TableRow({
                  height: { rule: 'atLeast', value: 500 },

                  children: [
                    new TableCell({
                      verticalAlign: 'center',
                      children: [
                        new Paragraph({
                          alignment: 'center',
                          children: [
                            new TextRun({
                              text: 'OVERALL RISK RATING',
                              bold: true,
                              size: 24,
                            }),
                          ],
                        }),
                      ],
                      borders,
                    }),
                    new TableCell({
                      verticalAlign: 'center',
                      children: [
                        new Paragraph({
                          alignment: 'center',
                          children: [
                            new TextRun({
                              text: data.risk_level,
                              bold: true,
                              size: 24,
                              color: getRiskColor(data.risk_level).color,
                            }),
                          ],
                        }),
                      ],

                      borders,
                      shading: {
                        fill: getRiskColor(data.risk_level).background, // Apply dynamic color
                      },
                    }),
                  ],
                }),
              ],
            }),
          ],
        },

        overall_summary: {
          type: PatchType.DOCUMENT,
          children: data.summary_of_findings
            .split(/\n+/)
            .map(
              (text) =>
                new Paragraph({ children: [new TextRun({ text, break: 1 })] }),
            ),
        },
        risk_areas: {
          type: PatchType.DOCUMENT,
          children: [
            new Table({
              width: {
                size: 75,
                type: WidthType.PERCENTAGE,
              },
              alignment: 'left',
              borders,
              rows: [
                new TableRow({
                  height: { rule: 'atLeast', value: 500 },
                  children: [
                    new TableCell({
                      verticalAlign: 'center',
                      width: {
                        size: 80,
                        type: WidthType.PERCENTAGE,
                      },

                      children: [
                        new Paragraph({
                          alignment: 'center',

                          children: [
                            new TextRun({
                              text: 'Risk Areas',
                              bold: true,
                              color: 'ffffff',
                            }),
                          ],
                        }),
                      ],
                      borders,
                      shading: {
                        fill: '595959', // Apply dynamic color
                      },
                    }),
                    new TableCell({
                      verticalAlign: 'center',
                      children: [
                        new Paragraph({
                          alignment: 'center',
                          children: [
                            new TextRun({
                              text: 'Risk Rating',
                              color: 'ffffff',
                              bold: true,
                            }),
                          ],
                        }),
                      ],
                      borders,
                      shading: {
                        fill: '595959', // Apply dynamic color
                      },
                    }),
                  ],
                }),
                ...data.riskData.map(
                  (risk) =>
                    new TableRow({
                      height: { rule: 'atLeast', value: 500 },
                      children: [
                        new TableCell({
                          verticalAlign: 'center',
                          children: [new Paragraph(risk.area)],
                          borders,
                        }),
                        new TableCell({
                          verticalAlign: 'center',
                          children: [
                            new Paragraph({
                              children: [
                                new TextRun({
                                  text: risk.rating,
                                  color: getRiskColor(risk.rating).color,
                                }),
                              ],
                              alignment: AlignmentType.CENTER,
                            }),
                          ],
                          shading: {
                            fill: getRiskColor(risk.rating).background,
                          },
                          borders,
                        }),
                      ],
                    }),
                ),
              ],
            }),
          ],
        },

        riskAreas_antiBriberyAndAntiCorruption: {
          type: PatchType.DOCUMENT,
          children: data.anti_summary.map((text) => {
            return new Paragraph({
              spacing: {
                before: 300,
                after: 300,
              },
              bullet: {
                level: 0,
              },

              children: [
                ...text
                  .trim()
                  .split(/\n+/)
                  .map(
                    (line, index) =>
                      new TextRun({
                        text: line,
                        break: index === 0 ? 0 : 1,
                      }),
                  ),

                new TextRun({
                  break: 1,
                }),
              ],
            });
          }),
        },

        ...Object.entries(data.riskAreas).reduce(
          (acc, [key, value]) => ({
            ...acc,
            [`riskAreas_${key}`]: {
              type: PatchType.DOCUMENT,
              children: value.map((text) => {
                return new Paragraph({
                  spacing: {
                    before: 300,
                    after: 300,
                  },
                  bullet: {
                    level: 0,
                  },

                  children: [
                    ...text
                      .trim()
                      .split(/\n+/)
                      .map(
                        (line, index) =>
                          new TextRun({
                            text: line,
                            break: index === 0 ? 0 : 1,
                          }),
                      ),

                    new TextRun({
                      break: 1,
                    }),
                  ],
                });
              }),
            },
          }),
          {},
        ),

        a_rating: highlightRating(data.riskData[0].rating),
        b_rating: highlightRating(data.riskData[1].rating),
        c_rating: highlightRating(data.riskData[2].rating),
        d_rating: highlightRating(data.riskData[3].rating),
        e_rating: highlightRating(data.riskData[4].rating),
        f_rating: highlightRating(data.riskData[5].rating),
        g_rating: highlightRating(data.riskData[6].rating),
        h_rating: highlightRating(data.riskData[7].rating),

        // Utils
        page_break: {
          type: PatchType.DOCUMENT,
          children: [new Paragraph({ pageBreakBefore: true })],
        },

        // Findings Content
        sanctions_findings: {
          type: PatchType.DOCUMENT,
          children: data.sanctions_findings
            ? data.sape_data.map(createFindingsTable).flat()
            : createNoHitsTable('SANCTIONS'),
        },
        pep_findings: {
          type: PatchType.DOCUMENT,
          children: data.pep_findings
            ? data.pep_data.map(createFindingsTable).flat()
            : createNoHitsTable('PeP'),
        },

        antiBribery_findings: {
          type: PatchType.DOCUMENT,
          children: data.bribery_findings
            ? data.bribery_data.map(createFindingsTable).flat()
            : createNoHitsTable('ANTI BRIBERY'),
        },
        antiCorruption_findings: {
          type: PatchType.DOCUMENT,
          children: data.corruption_findings
            ? data.corruption_data.map(createFindingsTable).flat()
            : createNoHitsTable('ANTI CORRUPTION'),
        },
        government_ownership_and_political_affiliations_findings: {
          type: PatchType.DOCUMENT,
          children: data.sown_findings
            ? data.sown_data.map(createFindingsTable).flat()
            : createNoHitsTable(
                'GOVERNMENT OWNERSHIP AND POLITICAL AFFILIATIONS',
              ),
        },
        financial_indicators_findings: {
          type: PatchType.DOCUMENT,
          children: data.financial_findings
            ? data.financial_data.map(createFindingsTable).flat()
            : createNoHitsTable('FINANCIALS'),
        },
        bankruptcy_findings: {
          type: PatchType.DOCUMENT,
          children: data.bankruptcy_findings
            ? data.backruptcy_data.map(createFindingsTable).flat()
            : createNoHitsTable('BANKRUPTCY'),
        },
        other_adverse_media_findings: {
          type: PatchType.DOCUMENT,
          children: data.adv_findings
            ? data.adv_data.map(createFindingsTable).flat()
            : createNoHitsTable('OTHER ADVERSE MEDIA'),
        },

        regularity_findings: {
          type: PatchType.DOCUMENT,
          children: data.reg_findings
            ? data.reg_data.map(createFindingsTable).flat()
            : createNoHitsTable('REGULATORY'),
        },
        legal_findings: {
          type: PatchType.DOCUMENT,
          children: data.bankruptcy_findings
            ? data.leg_data.map(createFindingsTable).flat()
            : createNoHitsTable('LEGAL'),
        },
        cyberSecurity_findings: {
          type: PatchType.DOCUMENT,
          children:
            data.cyberSecurity_findings.data.length > 0
              ? createFindingsInnerTable(data.cyberSecurity_findings)
              : createNoHitsTable('CYBER SECURITY'),
        },
        esg_findings: {
          type: PatchType.DOCUMENT,
          children:
            data.esg_findings.data.length > 0
              ? createFindingsInnerTable(data.esg_findings)
              : createNoHitsTable('ESG'),
        },
      },
    });

    const fileName = `${data.name}`;

    const docxPath = `src/${fileName}.docx`;
    const pdfPath = `src/${fileName}.pdf`;

    fs.writeFileSync(docxPath, doc);

    topdf.convert(docxPath, pdfPath);

    // Upload DOCX to Azure
    const docxUrl = await uploadToAzure(
      docxPath,
      `${data.ens_id}/${fileName}.docx`,
      data.session_id,
    );

    // Upload PDF to Azure
    const pdfUrl = await uploadToAzure(
      pdfPath,
      `${data.ens_id}/${fileName}.pdf`,
      data.session_id,
    );

    // Cleanup local files after upload
    fs.unlinkSync(docxPath);
    fs.unlinkSync(pdfPath);

    return { docxUrl, pdfUrl };
  } catch (error) {
    throw new Error(`Error generating report: ${error.message}`);
  }
};
