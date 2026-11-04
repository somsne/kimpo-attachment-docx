// 生成往返测试夹具（文档附件插件 kimpo-attachment-docx）：用 docx 包（MIT）生成 8 个覆盖常见结构的 docx，
// 写到 frontend/fixtures/roundtrip/。夹具入库，改了本脚本要重新运行：pnpm fixtures
//
// 覆盖：标题样式（含中文样式 ID 习惯）、编号与多级列表、表格（合并单元格、嵌套）、页眉页脚与页码、
// 图片、修订（插入 / 删除）、批注（含回复）、自定义段落样式与分页 / 分节。守门用的「不支持内容」样件
// 在单测里用 JSZip 现场拼，不在这里生成。
import {
  AlignmentType,
  Comment,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  InsertedTextRun,
  LevelFormat,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} from 'docx';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';

const outDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'fixtures', 'roundtrip');
mkdirSync(outDir, { recursive: true });

/** 修订与批注用固定时间（夹具内容稳定；zip 时间戳仍随生成时刻变化）。 */
const FIXED_DATE = new Date('2026-11-04T06:00:00Z');
const meta = { creator: 'Kimpo fixtures', lastModifiedBy: 'Kimpo fixtures', revision: 1 };

/** 生成 w×h 纯色 PNG（无第三方依赖）。 */
function solidPng(w, h, [r, g, b]) {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = buf => {
    let c = 0xffffffff;
    for (const byte of buf) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const sum = Buffer.alloc(4);
    sum.writeUInt32BE(crc(body));
    return Buffer.concat([len, body, sum]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const row = Buffer.concat([Buffer.from([0]), Buffer.from(Array.from({ length: w }, () => [r, g, b]).flat())]);
  const raw = Buffer.concat(Array.from({ length: h }, () => row));
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const cell = (text, opts = {}) => new TableCell({ children: [new Paragraph(text)], ...opts });

const fixtures = {
  '01-headings.docx': new Document({
    ...meta,
    sections: [
      {
        children: [
          new Paragraph({ text: '采购框架合同', heading: HeadingLevel.TITLE }),
          new Paragraph({ text: '一、合同标的', heading: HeadingLevel.HEADING_1 }),
          new Paragraph('甲方向乙方采购办公设备，具体型号见附件一。'),
          new Paragraph({ text: '二、价格与付款', heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: '2.1 单价', heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ children: [new TextRun('单价含税、含运费，'), new TextRun({ text: '合同期内不作调整', bold: true }), new TextRun('。')] }),
          new Paragraph({ text: '2.2 付款节点', heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: '2.2.1 首付款', heading: HeadingLevel.HEADING_3 }),
          new Paragraph({ children: [new TextRun({ text: 'Mixed English text with ', italics: true }), new TextRun({ text: 'underline', underline: {} }), new TextRun(' & symbols < > "quotes".')] })
        ]
      }
    ]
  }),
  '02-numbering.docx': new Document({
    ...meta,
    numbering: {
      config: [
        {
          reference: 'kimpo-multilevel',
          levels: [
            { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START },
            { level: 1, format: LevelFormat.DECIMAL, text: '%1.%2', alignment: AlignmentType.START },
            { level: 2, format: LevelFormat.LOWER_LETTER, text: '%3)', alignment: AlignmentType.START }
          ]
        }
      ]
    },
    sections: [
      {
        children: [
          new Paragraph({ text: '付款方式', heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ text: '首付款：合同签订后 10 日内支付 30%；', numbering: { reference: 'kimpo-multilevel', level: 0 } }),
          new Paragraph({ text: '到货款：每批到货验收合格后支付 60%；', numbering: { reference: 'kimpo-multilevel', level: 0 } }),
          new Paragraph({ text: '分批到货的按批次结算', numbering: { reference: 'kimpo-multilevel', level: 1 } }),
          new Paragraph({ text: '附验收单', numbering: { reference: 'kimpo-multilevel', level: 2 } }),
          new Paragraph({ text: '质保金：质保期满无质量问题后支付 10%。', numbering: { reference: 'kimpo-multilevel', level: 0 } }),
          new Paragraph({ text: '项目符号一', bullet: { level: 0 } }),
          new Paragraph({ text: '项目符号二', bullet: { level: 1 } })
        ]
      }
    ]
  }),
  '03-tables.docx': new Document({
    ...meta,
    sections: [
      {
        children: [
          new Paragraph('付款节点表：'),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({ tableHeader: true, children: ['付款节点', '比例', '金额（元）', '时限'].map(t => cell(t)) }),
              new TableRow({ children: ['首付款', '30%', '1,626,000.00', '签订后 10 日'].map(t => cell(t)) }),
              new TableRow({ children: [cell('到货款'), cell('60%'), cell('3,252,000.00', { columnSpan: 2 })] }),
              new TableRow({ children: [cell('质保金', { rowSpan: 2 }), cell('10%'), cell('542,000.00'), cell('质保期满')] }),
              new TableRow({ children: [cell('备注'), cell('—'), cell('—')] })
            ]
          }),
          new Paragraph('嵌套表格：'),
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph('外层单元格'),
                      new Table({ rows: [new TableRow({ children: [cell('内层 A'), cell('内层 B')] })] }),
                      new Paragraph('')
                    ]
                  }),
                  cell('外层右侧')
                ]
              })
            ]
          }),
          new Paragraph('表后段落。')
        ]
      }
    ]
  }),
  '04-header-footer.docx': new Document({
    ...meta,
    sections: [
      {
        headers: { default: new Header({ children: [new Paragraph('合同编号：CG-2026-011    机密 · 内部使用')] }) },
        footers: {
          default: new Footer({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun('第 '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' 页 / 共 '), new TextRun({ children: [PageNumber.TOTAL_PAGES] }), new TextRun(' 页')] })]
          })
        },
        children: [
          new Paragraph({ text: '页眉页脚样例', heading: HeadingLevel.HEADING_1 }),
          new Paragraph('正文第一页。'),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph('正文第二页。')
        ]
      }
    ]
  }),
  '05-images.docx': new Document({
    ...meta,
    sections: [
      {
        children: [
          new Paragraph('图片前的段落。'),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: 'png', data: solidPng(32, 20, [100, 108, 255]), transformation: { width: 160, height: 100 } })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, text: '图 1 交付流程' }),
          new Paragraph({ children: [new TextRun('行内图片：'), new ImageRun({ type: 'png', data: solidPng(8, 8, [250, 173, 20]), transformation: { width: 16, height: 16 } }), new TextRun(' 之后的文字。')] })
        ]
      }
    ]
  }),
  '06-revisions.docx': new Document({
    ...meta,
    features: { trackRevisions: true },
    sections: [
      {
        children: [
          new Paragraph({ text: '修订样例', heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new TextRun('合同期内单价'),
              new DeletedTextRun({ text: '不作调整', id: 1, author: '李四', date: FIXED_DATE.toISOString() }),
              new InsertedTextRun({ text: '每年复核一次，调整幅度不超过 3%', id: 2, author: '李四', date: FIXED_DATE.toISOString() }),
              new TextRun('。')
            ]
          }),
          new Paragraph('未修订的段落。')
        ]
      }
    ]
  }),
  '07-comments.docx': new Document({
    ...meta,
    comments: {
      children: [
        { id: 0, author: '张三', date: FIXED_DATE, children: [new Paragraph('付款期限财务要求改为 45 个工作日，请确认。')] },
        { id: 1, author: '李四', date: FIXED_DATE, children: [new Paragraph('已与财务确认。')] }
      ]
    },
    sections: [
      {
        children: [
          new Paragraph({ text: '批注样例', heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new TextRun('乙方开具合格发票后，甲方于 '),
              new CommentRangeStart(0),
              new TextRun('30 个工作日'),
              new CommentRangeEnd(0),
              new TextRun({ children: [new CommentReference(0)] }),
              new TextRun('内支付当期货款。')
            ]
          }),
          new Paragraph({ children: [new CommentRangeStart(1), new TextRun('第二条批注锚点'), new CommentRangeEnd(1), new TextRun({ children: [new CommentReference(1)] })] })
        ]
      }
    ]
  }),
  '08-styles-sections.docx': new Document({
    ...meta,
    styles: {
      paragraphStyles: [
        { id: 'ClauseBody', name: '条款正文', basedOn: 'Normal', quickFormat: true, run: { size: 22 }, paragraph: { indent: { firstLine: 480 }, spacing: { line: 360 } } },
        { id: '1', name: 'heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true }, paragraph: { outlineLevel: 0 } }
      ]
    },
    sections: [
      {
        children: [
          new Paragraph({ text: '中文样式 ID 的标题', style: '1' }),
          new Paragraph({ text: '自定义样式「条款正文」的段落。', style: 'ClauseBody' }),
          new Paragraph({ children: [new TextRun('分页前'), new PageBreak(), new TextRun('分页后')] })
        ]
      },
      {
        properties: { page: { size: { orientation: PageOrientation.LANDSCAPE } } },
        children: [new Paragraph({ text: '第二节（横向）', style: '1' }), new Paragraph({ text: '横向页面的条款正文。', style: 'ClauseBody' })]
      }
    ]
  })
};

for (const [name, doc] of Object.entries(fixtures)) {
  const buf = await Packer.toBuffer(doc);
  writeFileSync(join(outDir, name), buf);
  console.log(`${name}\t${buf.length} B`);
}
