import * as fs from "fs";
import pkg from 'docx';
const { Document, Packer, WidthType, convertInchesToTwip, VerticalAlign, AlignmentType, TextRun, ShadingType, VerticalPositionAlign, HorizontalPositionAlign, FrameAnchorType, Paragraph, BorderStyle, PageBorderDisplay, Table, TableCell, TableRow, PageBorderOffsetFrom, PageBorderZOrder } = pkg;

const doc = new Document({
    sections: [
        {
            properties: {
                // page: {
                //     borders: {
                //         pageBorderBottom: {
                //             style: BorderStyle.SINGLE,
                //             size: 2 * 8, //2pt;
                //             color: "000000",
                //         },
                //         pageBorderLeft: {
                //             style: BorderStyle.SINGLE,
                //             size: 1 * 8, //1pt;
                //             color: "000000",
                //         },
                //         pageBorderRight: {
                //             style: BorderStyle.SINGLE,
                //             size: 1 * 8, //1pt;
                //             color: "FF00AA",
                //         },
                //         pageBorderTop: {
                //             style: BorderStyle.SINGLE,
                //             size: 1 * 8, //1pt;
                //             color: "000000",
                //         },
                //         pageBorders: {
                //             display: PageBorderDisplay.ALL_PAGES,
                //             offsetFrom: PageBorderOffsetFrom.TEXT,
                //             zOrder: PageBorderZOrder.FRONT,
                //         },
                //     },
                // },
            },
            children: [
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ' Design Doc - Program and its Entities',
                                                    bold: true,
                                                    // highlight: "blue",
                                                    underline: {}
                                                }),
                                            ],
                                        }),
                                    ],
                                    shading: {
                                        fill: "D3DEDC",
                                        type: ShadingType.CLEAR,
                                        color: "D3DEDC",
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                    },    
                                }),
                            ],
                        }),
                    ],
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ` Program Name:`,
                                                    bold: true,
                                                }),
                                            ],
                                        }),
                                    ],
                                    shading: {
                                        fill: "D3DEDC",
                                        type: ShadingType.CLEAR,
                                        color: "D3DEDC",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                    },       
                                }),
                            ],
                        }),
                    ],
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `CONTCTINQ`,
                            bold: true
                        }),
                    ],
                    alignment: AlignmentType.CENTER,
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ` Entities Used:`,
                                                    bold: true,
                                                }),
                                            ],
                                        }),
                                    ],
                                    shading: {
                                        fill: "D3DEDC",
                                        type: ShadingType.CLEAR,
                                        color: "D3DEDC",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "D3DEDC",
                                        },
                                    },       
                                }),
                            ],
                        }),
                    ],
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ' Not Updated',
                                                    bold: true
                                                }),
                                            ],
                                        }),
                                    ],
                                    width: {
                                        size: 33,
                                        type: WidthType.PERCENTAGE,
                                    },
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },        
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ' AGENTS',
                                                    bold: true
                                                }),
                                            ],
                                        }),
                                    ],
                                    width: {
                                        size: 33,
                                        type: WidthType.PERCENTAGE,
                                    },
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },           
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ' agents',
                                                    bold: true
                                                }),
                                            ],
                                        }),
                                    ],
                                    width: {
                                        size: 33,
                                        type: WidthType.PERCENTAGE,
                                    },
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },   
                                    },           
                                }),
                            ],
                        }),
                    ],
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Table({
                    alignment: AlignmentType.RIGHT,
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph(" Not Updated\t")
                                    ],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },  
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },        
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph(" AID\t\t\t")
                                    ],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },          
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph("\t5 a\t")
                                    ],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                    },         
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tAgent ID\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },          
                                }),
                            ],
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ' Updated',
                                                    bold: true
                                                }),
                                            ],
                                            shading: {
                                                fill: "F4EEFF",
                                                type: ShadingType.CLEAR,
                                                color: "DCD6F7",
                                            },    
                                        }),
                                    ],
                                    width: {
                                        size: 33,
                                        type: WidthType.PERCENTAGE,
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },      
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ' CONTACT',
                                                    bold: true
                                                }),
                                            ],
                                            shading: {
                                                fill: "F4EEFF",
                                                type: ShadingType.CLEAR,
                                                color: "DCD6F7",
                                            },          
                                        }),
                                    ],
                                    width: {
                                        size: 33,
                                        type: WidthType.PERCENTAGE,
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                    },
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: ' Contact',
                                                    bold: true
                                                }),
                                            ],
                                            shading: {
                                                fill: "F4EEFF",
                                                type: ShadingType.CLEAR,
                                                color: "DCD6F7",
                                            },     
                                        }),
                                    ],
                                    width: {
                                        size: 33,
                                        type: WidthType.PERCENTAGE,
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },     
                                }),
                            ],
                        }),
                    ],
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `\n`,
                        }),
                    ],
                }),
                new Table({
                    alignment: AlignmentType.RIGHT,
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph(" Not Updated\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },  
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },        
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tCCCCID\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },       
                                }),
                                new TableCell({
                                    children: [new Paragraph("\t5 a\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },         
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tID\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },           
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [ new Paragraph(" Updated\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },  
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },          
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tCCCUID\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },        
                                }),
                                new TableCell({
                                    children: [new Paragraph("\t5 a\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },         
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tCust ID\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },  
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },          
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [ new Paragraph(" Not Updated\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },           
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tCCNAME\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },         
                                }),
                                new TableCell({
                                    children: [new Paragraph("\t2p0\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },  
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },       
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tName\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },  
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },          
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [ new Paragraph(" Updated\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },           
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tCCSEQ\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },        
                                }),
                                new TableCell({
                                    children: [new Paragraph("\t5 a\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                    },        
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tConst Seq\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    },  
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        // bottom: {
                                        //     style: BorderStyle.CLEAR,
                                        //     size: 1,
                                        //     color: "F4EEFF",
                                        // },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },          
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [ new Paragraph(" Not Updated\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        left: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },           
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tCCSTSID\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                    },        
                                }),
                                new TableCell({
                                    children: [new Paragraph("\t3p3\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                    },        
                                }),
                                new TableCell({
                                    children: [new Paragraph("\tStatus\t")],
                                    shading: {
                                        fill: "F4EEFF",
                                        type: ShadingType.CLEAR,
                                        color: "DCD6F7",
                                    }, 
                                    borders: {
                                        top: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        bottom: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        },
                                        right: {
                                            style: BorderStyle.CLEAR,
                                            size: 1,
                                            color: "F4EEFF",
                                        }, 
                                    },           
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        },
    ],
});


// Used to export the file into a .docx file
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});