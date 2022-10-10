import {
    saveAs
} from './file-saver';
import * as docx from 'docx';

let str = []; // 用于存储网页的信息，图片、文字、表格等

/**
 * htmlExportWorde函数的参数说明
 * domName 传入的dom元素的类名
 * fileName 导出的文件名
 */
export function htmlExportWord(domName, fileName) {
    const doc = new docx.Document();
    let domStr = []; // 存放第一次获取到的页面信息，此时有可能为多维数组
    let arr = []; // 把第一次获取的页面信息全部展开存放到arr数组里面
    let imageBuffer = null; // 用于存放图片转换后的信息
    let imgBuffer = []; // 处理图片后的页面信息

    let classNames = document.getElementsByClassName(domName)[0];
    domStr.push(resolveDom(classNames)); // 获取页面的信息、标签名称等
    // str为多维数组，所以需要展开重新放到数组里面
    domStr.forEach(v => {
        arr.push(...v);
    })
    arr.forEach(v => {
        if (v.tagNames == 'IMG') { // 转换图片
            imageBuffer = fetch(v.srcOrText).then((response) => response.arrayBuffer());
            imgBuffer.push({
                tagName: 'IMG',
                text: docx.Media.addImage(doc, imageBuffer, v.imgWidth, v.imgHeight)
            });
        } else {
            imgBuffer.push(v);
        }
    })

    doc.addSection({
        children: [
            ...(function () {
                let arrs = [];
                imgBuffer.forEach(v => {
                    if (v.tagName == 'TABLE') { // 表格
                        arrs.push(
                            new docx.Table({
                                width: {
                                    size: 5000,
                                    type: docx.WidthType.PERCENTAGE, // 导出的表格宽度percentage
                                },
                                rows: [...createDocxTableRow(v.children)],
                                spacing: {
                                    before: 200, // 段落间距/前
                                },
                            })
                        )
                    } else if (v.tagName == 'IMG') { // 图片
                        arrs.push(
                            new docx.Paragraph({
                                children: [v.text],
                                alignment: docx.AlignmentType.CENTER, // 水平居中
                            })
                        )
                    } else { // 文本等
                        if (v.tagId == 'head-line') { // 大标题
                            arrs.push(
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: v.srcOrText,
                                            size: 36, // 字体大小
                                            bold: true, // 加粗
                                        }),
                                    ],
                                    spacing: {
                                        after: 200, // 段落间距/后
                                    },
                                    alignment: docx.AlignmentType.CENTER, // 文本居中
                                })
                            )
                        } else if (/sub-title/.test(v.classNames)) { // 小标题
                            arrs.push(
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: v.srcOrText,
                                            size: 24,
                                            bold: true, // 加粗
                                        }),
                                    ],
                                    spacing: {
                                        after: 200, // 段落间距/后
                                    },
                                })
                            )
                        } else { // 其他
                            arrs.push(
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: v.srcOrText,
                                            // bold: true, // 加粗
                                        }),
                                    ],
                                    spacing: {
                                        after: 200, // 段落间距/后
                                    },
                                    indent: {
                                        firstLine: 500 // 首行缩进
                                    }
                                })
                            )
                        }
                    }
                })
                return arrs
            })(),
        ],
    });
    docx.Packer.toBlob(doc).then((blob) => {
        saveAs(blob, fileName + ".docx");
        console.log("文档生成成功");
        str = []; // 每次导出后清空数组
    });
    // 导出后清空，防止累加
    doc.addSection = () => {}
}

// 设置表格行
function createDocxTableRow(v) {
    let tableRowArr = []
    v.forEach(item => {
        let firstRow = null;
        if (item.children.length > 0) {
            firstRow = item.children[0];
            tableRowArr.push(
                new docx.TableRow({
                    children: [
                        ...createDocxTableCell(item.children),
                    ],
                }),
            );
        }
    });
    return tableRowArr;
}

// 设置表格的单元格
function createDocxTableCell(item) {
    let tableCellArr = [];
    item.forEach(element => {
        tableCellArr.push(
            new docx.TableCell({
                children: [new docx.Paragraph({
                    text: element.text,
                    alignment: docx.AlignmentType.CENTER, // 文本水平居中
                })],
                cantSplit: true, // 防止分页
                verticalAlign: docx.VerticalAlign.CENTER, // 文本对齐方式
                rowSpan: element.rowspan,
                columnSpan: element.colspan,
            }),
        )
    })
    return tableCellArr;
}


/**
 * tableObj数据格式
 *  tableObj = {
 *  tagName: 'TABLE',
 *    children: [{
 *      tagName: 'th',
 *      children: [{text:'1', rowspan, colspan}, {text:'2', rowspan, colspan}, {text:'3', rowspan, colspan}]
 *    },
 *    {
 *      tagName: 'td',
 *      children: [{text:'1', rowspan, colspan}, {text:'2', rowspan, colspan}, {text:'3', rowspan, colspan}]
 *    }]
 *  }
 */

/**
 * 获取页面显示的数据，文本、表格、图片等信息,canvas需要先转换base64
 */
function resolveDom(dom) {
    if (dom.tagName == 'TABLE') { // 表格的处理
        let tableObj = {
            tagName: 'TABLE',
            children: []
        };
        Array.from(dom.children).forEach(v => {
            let domTagName = v.lastChild.tagName // 标签名称
            let domArr = [];
            Array.from(v.children).forEach(item => {
                domArr.push({
                    text: item.innerText,
                    colspan: item.getAttribute('colspan'),
                    rowspan: item.getAttribute('rowspan'),
                })
            })
            tableObj.children.push({
                tagName: domTagName,
                children: domArr
            })
        })
        str.push(tableObj);
    } else if (dom.children.length != 0) { // 是否还有子标签
        // 有子标签但是也不再进行拆分，直接看作整体获取文本信息，可根据需求自行修改条件
        if (/span-arr/.test(dom.getAttribute('class'))) {
            str.push({
                tagNames: dom.tagName,
                srcOrText: dom.innerText,
                classNames: dom.className,
            })
        } else { // 有子标签，继续拆分
            for (let i = 0; i < dom.children.length; i++) {
                resolveDom(dom.children[i])
            }
        }
    } else if (dom.tagName == 'IMG') { // 图片的处理
        if (dom.width > 900) { // 图片宽度大于900px，导出文档时设置统一宽高，否则获取自身宽高赋值
            str.push({
                tagNames: dom.tagName,
                srcOrText: dom.src,
                imgWidth: 700,
                imgHeight: 350,
            })
        } else {
            str.push({
                tagNames: dom.tagName,
                srcOrText: dom.src,
                imgWidth: dom.width,
                imgHeight: dom.height,
            })
        }
    } else { // 其余标签的处理
        str.push({
            tagNames: dom.tagName,
            srcOrText: dom.innerText,
            tagId: dom.getAttribute('id'),
            classNames: dom.className,
        })
    }
    return str;
}