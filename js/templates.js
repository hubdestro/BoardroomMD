/**
 * The Template Factory Configuration
 * Defines presentation styles separate from logic.
 */

window.TemplateFactory = {
    corporate: {
        creator: "BoardroomMD",
        styles: {
            default: {
                document: {
                    run: {
                        font: "Arial",
                        size: 24, // 12pt
                        color: "000000"
                    },
                    paragraph: {
                        spacing: { line: 276, before: 200, after: 200 }
                    }
                }
            },
            paragraphStyles: [
                {
                    id: "CoverTitle",
                    name: "Cover Title",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 72, bold: true, color: "000000" }, // 36pt
                    paragraph: { spacing: { before: 2000, after: 400 }, alignment: window.docx.AlignmentType.CENTER }
                },
                {
                    id: "CoverSubtitle",
                    name: "Cover Subtitle",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 36, italic: true, color: "555555" }, // 18pt
                    paragraph: { spacing: { before: 200, after: 200 }, alignment: window.docx.AlignmentType.CENTER }
                },
                {
                    id: "Heading1",
                    name: "Heading 1",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 36, bold: true, color: "0F4C81" }, // Classic Blue
                    paragraph: { spacing: { before: 400, after: 200 } }
                },
                {
                    id: "Heading2",
                    name: "Heading 2",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 32, bold: true, color: "333333" },
                    paragraph: { spacing: { before: 300, after: 150 } }
                },
                {
                    id: "Quote",
                    name: "Quote",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { italic: true, color: "666666" },
                    paragraph: { indent: { left: 720 }, borders: { left: { color: "0F4C81", space: 10, size: 20, style: "single" } } }
                }
            ]
        },
        getSections: function(docxElements, metadata) {
            const { Paragraph, TextRun, Header, Footer, AlignmentType, PageNumber, SectionType } = window.docx;
            
            // Cover Page Elements
            const coverPage = [
                new Paragraph({ text: metadata.title || "Document Title", style: "CoverTitle" }),
                new Paragraph({ text: metadata.author || "Author", style: "CoverSubtitle" }),
                new Paragraph({ text: new Date().toLocaleDateString(), style: "CoverSubtitle" }),
                new Paragraph({ children: [new TextRun({ text: "", break: 1 })], pageBreakBefore: true }) // Force page break after cover
            ];

            return [
                {
                    properties: {
                        type: SectionType.NEXT_PAGE,
                    },
                    headers: {
                        default: new Header({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new TextRun({
                                            text: metadata.title || "BoardroomMD Document",
                                            color: "999999",
                                            italics: true
                                        })
                                    ]
                                })
                            ]
                        })
                    },
                    footers: {
                        default: new Footer({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                        new TextRun("Page "),
                                        new TextRun({ children: [PageNumber.CURRENT] }),
                                        new TextRun(" of "),
                                        new TextRun({ children: [PageNumber.TOTAL_PAGES] })
                                    ]
                                })
                            ]
                        })
                    },
                    children: [...coverPage, ...docxElements]
                }
            ];
        }
    },

    modern: {
        creator: "BoardroomMD",
        styles: {
            default: {
                document: {
                    run: {
                        font: "Helvetica",
                        size: 22, // 11pt
                        color: "111111"
                    },
                    paragraph: {
                        spacing: { line: 360, before: 300, after: 300 }
                    }
                }
            },
            paragraphStyles: [
                {
                    id: "CoverTitle",
                    name: "Cover Title",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 96, bold: true, color: "22c55e" }, // Brand color, massive size
                    paragraph: { spacing: { before: 3000, after: 600 }, alignment: window.docx.AlignmentType.LEFT }
                },
                {
                    id: "CoverSubtitle",
                    name: "Cover Subtitle",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 40, color: "555555", bold: false }, 
                    paragraph: { spacing: { before: 200, after: 200 }, alignment: window.docx.AlignmentType.LEFT }
                },
                {
                    id: "Heading1",
                    name: "Heading 1",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 48, bold: false, color: "22c55e" }, // Brand color
                    paragraph: { spacing: { before: 600, after: 300 } }
                },
                {
                    id: "Heading2",
                    name: "Heading 2",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 36, bold: false, color: "333333" },
                    paragraph: { spacing: { before: 400, after: 200 } }
                },
                {
                    id: "Quote",
                    name: "Quote",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { italic: true, color: "888888" },
                    paragraph: { 
                        indent: { left: 720 },
                        borders: { left: { color: "22c55e", space: 10, size: 20, style: "single" } }
                    } 
                }
            ]
        },
        getSections: function(docxElements, metadata) {
            const { Paragraph, TextRun, Header, Footer, AlignmentType, PageNumber, SectionType, BorderStyle } = window.docx;
            
            const coverPage = [
                new Paragraph({ text: metadata.title || "Document Title", style: "CoverTitle" }),
                new Paragraph({ text: metadata.author || "Author", style: "CoverSubtitle" }),
                new Paragraph({ text: new Date().toLocaleDateString(), style: "CoverSubtitle" }),
                new Paragraph({ children: [new TextRun({ text: "", break: 1 })], pageBreakBefore: true })
            ];

            return [
                {
                    properties: {
                        type: SectionType.NEXT_PAGE,
                    },
                    headers: {
                        default: new Header({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    border: { bottom: { color: "22c55e", space: 1, style: BorderStyle.SINGLE, size: 6 } },
                                    children: [
                                        new TextRun({
                                            text: (metadata.title || "BoardroomMD").toUpperCase(),
                                            color: "22c55e",
                                            bold: true,
                                            size: 20
                                        })
                                    ]
                                })
                            ]
                        })
                    },
                    footers: {
                        default: new Footer({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new TextRun({
                                            children: ["Page ", PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES],
                                            color: "aaaaaa",
                                            size: 18
                                        })
                                    ]
                                })
                            ]
                        })
                    },
                    children: [...coverPage, ...docxElements]
                }
            ];
        }
    },

    // ──────────────────────────────────────────────────────────────────────────
    // PREMIUM TEMPLATE  (mirrors Document-t-1 & Document-t-2 reference files)
    // Cover:  full-page gradient PNG extracted from Document-t-1.docx
    // Body:   dark olive (#454F3C) text on soft sage (#e7ebe4) background
    // ──────────────────────────────────────────────────────────────────────────
    premium: {
        creator: "BoardroomMD",
        background: { color: "e7ebe4" },
        styles: {
            default: {
                document: {
                    run: {
                        font: "Calibri",
                        size: 22,
                        color: "454F3C"
                    },
                    paragraph: {
                        spacing: { line: 300, before: 160, after: 160 }
                    }
                }
            },
            paragraphStyles: [
                {
                    id: "CoverTitle",
                    name: "Cover Title",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 80, bold: true, color: "FFFFFF" },
                    paragraph: { spacing: { before: 4000, after: 400 }, alignment: window.docx.AlignmentType.LEFT }
                },
                {
                    id: "CoverSubtitle",
                    name: "Cover Subtitle",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 36, color: "DDDDDD", bold: false },
                    paragraph: { spacing: { before: 160, after: 160 }, alignment: window.docx.AlignmentType.LEFT }
                },
                {
                    id: "Heading1",
                    name: "Heading 1",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 40, bold: false, color: "454F3C" },
                    paragraph: {
                        spacing: { before: 400, after: 200 },
                        border: { bottom: { color: "677659", space: 4, size: 8, style: "single" } }
                    }
                },
                {
                    id: "Heading2",
                    name: "Heading 2",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { size: 32, bold: true, color: "677659" },
                    paragraph: { spacing: { before: 300, after: 150 } }
                },
                {
                    id: "Quote",
                    name: "Quote",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: { italic: true, color: "677659" },
                    paragraph: {
                        indent: { left: 720 },
                        borders: { left: { color: "677659", space: 10, size: 20, style: "single" } }
                    }
                }
            ]
        },
        getSections: function(docxElements, metadata) {
            const {
                Paragraph, TextRun, ImageRun, Header, Footer,
                AlignmentType, PageNumber, SectionType, BorderStyle,
                HorizontalPositionAlign, VerticalPositionAlign,
                HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom
            } = window.docx;

            function base64ToBytes(b64) {
                const raw = atob(b64);
                const arr = new Uint8Array(raw.length);
                for (let i = 0; i < raw.length; i++) arr[i] = raw.charCodeAt(i);
                return arr;
            }

            const imgBytes = base64ToBytes(window.AppAssets.coverGradient);

            // Cover Image anchored to page edges
            const bgImage = new ImageRun({
                data: imgBytes,
                transformation: { width: 793, height: 1122 },
                floating: {
                    horizontalPosition: {
                        relative: HorizontalPositionRelativeFrom.PAGE,
                        align: HorizontalPositionAlign.LEFT
                    },
                    verticalPosition: {
                        relative: VerticalPositionRelativeFrom.PAGE,
                        align: VerticalPositionAlign.TOP
                    },
                    behindDocument: true,
                    allowOverlap: true
                }
            });

            // ─── Section 1: COVER PAGE ONLY ──────────────────────────────────
            const coverSection = {
                properties: {
                    type: SectionType.NEXT_PAGE,
                    page: {
                        margin: { top: 0, right: 0, bottom: 0, left: 0 } // Zero margins for full-bleed cover
                    }
                },
                children: [
                    new Paragraph({ children: [bgImage] }),
                    new Paragraph({ text: metadata.title || "Document Title", style: "CoverTitle" }),
                    new Paragraph({ text: metadata.author || "Author", style: "CoverSubtitle" }),
                    new Paragraph({ text: new Date().toLocaleDateString(), style: "CoverSubtitle" })
                ]
            };

            // ─── Section 2: BODY CONTENT ─────────────────────────────────────
            // We clone docxElements and add a page break to the first element 
            // as a fail-safe for pagination in some previews.
            const bodyChildren = [...docxElements];
            if (bodyChildren.length > 0) {
                // Wrap the first element to ensure it has a page break if it's a paragraph
                // or just prepend a break.
                bodyChildren.unshift(new Paragraph({ text: "", pageBreakBefore: true }));
            }

            const bodySection = {
                properties: {
                    type: SectionType.NEXT_PAGE,
                    page: {
                        margin: {
                            top: 1440,    // 1 inch
                            right: 1440,  // 1 inch
                            bottom: 1440, // 1 inch
                            left: 1440    // 1 inch
                        }
                    }
                },
                headers: {
                    default: new Header({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.RIGHT,
                                border: { bottom: { color: "677659", space: 1, style: BorderStyle.SINGLE, size: 6 } },
                                children: [
                                    new TextRun({ text: metadata.title || "BoardroomMD", color: "454F3C", size: 20 })
                                ]
                            })
                        ]
                    })
                },
                footers: {
                    default: new Footer({
                        children: [
                            new window.docx.Table({
                                width: { size: 100, type: window.docx.WidthType.PERCENTAGE },
                                borders: {
                                    top: window.docx.TableBorders.NONE,
                                    bottom: window.docx.TableBorders.NONE,
                                    left: window.docx.TableBorders.NONE,
                                    right: window.docx.TableBorders.NONE,
                                    insideHorizontal: window.docx.TableBorders.NONE,
                                    insideVertical: window.docx.TableBorders.NONE,
                                },
                                rows: [
                                    new window.docx.TableRow({
                                        children: [
                                            new window.docx.TableCell({
                                                width: { size: 33, type: window.docx.WidthType.PERCENTAGE },
                                                children: [],
                                                borders: window.docx.TableBorders.NONE,
                                            }),
                                            new window.docx.TableCell({
                                                width: { size: 33, type: window.docx.WidthType.PERCENTAGE },
                                                children: [
                                                    new Paragraph({
                                                        alignment: AlignmentType.CENTER,
                                                        children: [
                                                            new TextRun({ text: "BoardroomMD", color: "677659", size: 18 })
                                                        ]
                                                    })
                                                ],
                                                borders: window.docx.TableBorders.NONE,
                                            }),
                                            new window.docx.TableCell({
                                                width: { size: 33, type: window.docx.WidthType.PERCENTAGE },
                                                children: [
                                                    new Paragraph({
                                                        alignment: AlignmentType.RIGHT,
                                                        children: [
                                                            new TextRun({ text: "Page ", color: "677659", size: 18 }),
                                                            new TextRun({ children: [PageNumber.CURRENT], color: "677659", size: 18 }),
                                                            new TextRun({ text: " of ", color: "677659", size: 18 }),
                                                            new TextRun({ children: [PageNumber.TOTAL_PAGES], color: "677659", size: 18 })
                                                        ]
                                                    })
                                                ],
                                                borders: window.docx.TableBorders.NONE,
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ]
                    })
                },
                children: bodyChildren
            };

            return [coverSection, bodySection];
        }
    }
};



