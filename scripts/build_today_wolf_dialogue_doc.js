const fs = require("fs");
const path = require("path");
const Module = require("module");

process.env.NODE_PATH = [
  process.env.NODE_PATH || "",
  "C:/Users/wang_/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules",
].filter(Boolean).join(path.delimiter);
Module._initPaths();

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  WidthType,
  BorderStyle,
  ShadingType,
  Footer,
  PageNumber,
  LevelFormat,
} = require("docx");

const outDir = path.resolve("outputs");
fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, "狼大今日发言与对话分析整理_20260701.docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "D9DEE8" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 90, bottom: 90, left: 120, right: 120 };

function t(text, opts = {}) {
  return new TextRun({ text: String(text), font: "Microsoft YaHei", size: opts.size || 22, bold: opts.bold, italics: opts.italics, color: opts.color });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 60, after: opts.after || 80, line: 320 },
    alignment: opts.alignment,
    children: Array.isArray(text) ? text : [t(text, opts)],
  });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 260, after: 180 }, children: [t(text, { bold: true, size: 32 })] });
}

function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 120 }, children: [t(text, { bold: true, size: 27 })] });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 70, line: 300 },
    children: [t(text)],
  });
}

function cell(text, width, shade = false, bold = false) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { fill: "EAF2F8", type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [p(text, { after: 0, bold })],
  });
}

function table(headers, rows, widths) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((x, i) => cell(x, widths[i], true, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((x, i) => cell(x, widths[i])) })),
    ],
  });
}

const wolfRows = [
  ["3494", "聚和材料", "核心不只是掩膜版映射，而是收购韩方技术能否在国内落地，厂子已建好，处于技术转移阶段。", "看技术/产能落地，不只看概念标签。"],
  ["3496", "天赐材料", "公司本身没问题，但新能源弱一些，且狼大明确说不在这个。", "可观察但降权，不能据此猜持仓。"],
  ["3497/3502", "电池/储能", "重新开新能源，是亿纬锂能业绩后才找到逻辑确认；AIDC需要储能扩充，储能加速且利润未被上游长协明显侵蚀。", "电池方向先看业绩确认，再看资金外溢与筹码清洗。"],
  ["3504", "半导体材料", "设备最多上修，材料消耗随投片和扩产放大，国产替代扩产是核心。", "下半年材料看持续消耗，不把设备逻辑粗暴套到材料。"],
  ["3506", "电网", "国内电网瓶颈与海外不同，增量来自AIDC加速建设和IDC转AIDC改造。", "电网不是泛缺电故事，要看AIDC改造增量。"],
  ["3510", "AIDC", "围绕下半年AIDC，狼大回答方向是Token工厂。", "重点从普通IDC/算力租赁转向持续Token消耗与华为国算链。"],
  ["3511", "指数与3-5", "指数并没有很好；临近中报，白线方向很多钱想跑；3-5可能是缩圈、再缩圈的加速行情。", "行情不是普涨一步到位，要防第二只脚和中报期指数破位。"],
  ["3512", "小金属/稀土", "小金属没什么问题，稀土小金属也是半导体材料；但属于小作文加速型。", "小作文加速型按小波段，不等同于趋势核心沿线持有。"],
  ["3513/3516", "被动止盈", "进入四浪的判断，落到手上一半盈利仓触发被动止盈；非长线持仓分散到近20只以降低中报黑天鹅。", "用账户触发条件替代主观猜顶；破了变现金，再找新逻辑。"],
  ["3517", "整车出海", "整车出海故事存在，但利好逻辑抵不过消费疲软。", "新能源内部也要分清电池/储能与整车消费。"],
  ["3518", "材料利润压制", "除非重启关税战，否则暂不需要担心材料和小金属利润被额外使命压制。", "材料逻辑仍回到产业需求、涨价、订单和量价。"],
  ["3519", "术与信息", "对狼大来说万物皆周期，技术框架可套白酒、新能源、医药、科技；吃超额才需要更多消息。", "普通人优先打磨技术框架、资金理解、盘口和心态。"],
  ["3520", "智驾", "L3/L4标准催化未动，关键在法律法规这个质变门槛。", "政策文本不等于产业兑现，需看真正商业化约束是否解除。"],
];

const questionRows = [
  ["新能源/聚和", "狼大最近做新能源，已提聚和，按他的思路还有哪些票可观察？", "以聚和拆模板：技术/产能落地、低位或支撑、业绩确认、储能/AIDC传导；结合最新发言，整车出海降权，电池/储能仍看业绩。", "优先观察科达利、星源材质、鹏辉能源、盛弘股份；嘉元科技、新宙邦、多氟多、麦格米特逻辑强但注意加速。"],
  ["半导体材料/石英", "已知狼大进石英股份，半导体材料里国产替代和卡位稀缺有哪些？", "先把石英作为高纯石英材料锚，再按硅片、掩膜版、CMP、前驱体、光刻胶、靶材、特气分层；最新发言进一步区分趋势核心与小作文加速型小金属。", "更值得深挖菲利华、路维光电、沪硅产业、立昂微、石英股份；核心但等回踩：鼎龙、安集、雅克、南大、彤程、江丰。小金属按小波段处理。"],
  ["AIDC/机构拿筹码", "AIDC方向，符合狼大思路且近期成交量放大、疑似机构拿筹码的是哪些？", "用腾讯财经复权日线补5日/20日、10日/60日量比；东方财富资金流接口本次不可用，因此不把“主力流入”写成事实。", "第一观察池保留科华数据、英维克、华丰科技、同飞股份、申菱环境、中科曙光、浪潮信息、瑞可达；结论降级为“量价疑似有大资金活动，需龙虎榜/公告/调研验证”。"],
];

const newEnergyRows = [
  ["科达利 002850", "电池结构件", "6月涨幅约8.9%，站上5/20/144日线，接近60日线。", "产业链传导清晰，观察中报、客户订单和60日线修复。"],
  ["星源材质 300568", "隔膜", "6月涨幅约18.3%，站上主要均线。", "隔膜价格/盈利修复若有公告或研报支撑，可继续跟踪。"],
  ["鹏辉能源 300438", "储能电池", "6月涨幅约8.3%，站上主要均线。", "更直接受益储能/AIDC，但需业绩验证。"],
  ["盛弘股份 300693", "储能PCS/电力电子", "6月仍为负涨幅，站上5/20/144但未站回60日线。", "适合看右侧修复和储能订单确认。"],
];

const materialRows = [
  ["石英股份 603688", "高纯石英/半导体石英材料", "已知锚点，6月涨约23%，年内涨幅较大。", "看5日线修复、半导体石英订单和量价承接。"],
  ["菲利华 300395", "石英玻璃材料/石英电子布", "6月基本未明显加速。", "与石英不是同赛道，关注石英电子布认证和订单。"],
  ["路维光电 688401", "半导体掩膜版", "6月涨约22%，较温和。", "看半导体掩膜版项目与客户导入。"],
  ["沪硅产业 688126", "12英寸大硅片", "国产大硅片稀缺，但仍有盈利压力。", "看亏损收窄、300mm出货和价格。"],
  ["立昂微 605358", "硅片+功率半导", "主板硅片代餐属性。", "看硅片收入占比与功率半导拖累程度。"],
];

const aidcRows = [
  ["科华数据 002335", "UPS/数据中心/储能", "腾讯财经：5日/20日量=1.59，10日/60日量=1.37，6月+8.7%。", "最像“悄悄放量、未明显加速”的电源配套；资金流待补。"],
  ["英维克 002837", "液冷/温控", "腾讯财经：5日/20日量=1.22，10日/60日量=1.40，6月+13.9%。", "符合中报后验证的液冷链，位置未过热；看中报和华为链证据。"],
  ["华丰科技 688629", "华为国算链/高速连接器", "腾讯财经：5日/20日量=1.23，6月+32.1%；狼大明示锚点。", "证据强，但已被市场看见，不是刚启动。"],
  ["同飞股份 300990", "液冷温控", "腾讯财经：5日/20日量=1.19，10日/60日量=1.32，6月+29.7%。", "强度起来，需看回踩承接；资金流待补。"],
  ["申菱环境 301018", "温控/液冷", "腾讯财经：5日/20日量=1.16，10日/60日量=1.42，6月+13.1%。", "量能形态好，但年内涨幅较大。"],
  ["中科曙光 603019", "国产服务器/算力基建", "腾讯财经：5日/20日量=1.90，10日/60日量=1.35，6月+22.1%。", "更像机构指数级资金，非纯华为卡逻辑。"],
  ["浪潮信息 000977", "AI服务器", "腾讯财经：5日/20日量=1.24，6月+8.3%。", "服务器大票量价改善，偏基础设施锚。"],
  ["瑞可达 688800", "高速连接器", "腾讯财经：5日/20日量=1.14，10日/60日量=1.27，6月+12.7%。", "连接器链补充观察。"],
];

const dataRows = [
  ["论坛抓取", "成功增量更新18条，本周文档共198条；全库最新发言2026-06-30 23:40，狼大最新2026-06-30 22:33。", "3个用户分页触发403，已保留成功用户增量。"],
  ["行情数据", "东方财富指数接口在本周文档生成时断连，脚本已回落腾讯财经；0701专题候选量价使用腾讯财经复权日线。", "后续固定规则：东方财富失败即用腾讯财经。"],
  ["资金流", "东方财富历史资金流接口本轮断连/限流，未能取得稳定主力净流入数据。", "不把“机构拿筹码”写成事实，只写量价疑似并等待龙虎榜、基金持仓或资金流恢复验证。"],
  ["公告/研报", "本次将公告、研报、调研纪要作为验证项写入，而非把未核实小作文作为结论。", "后续针对候选票补公司公告、业绩预告、研报页和调研纪要。"],
];

const children = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [t("狼大今日发言与本日对话分析整理", { bold: true, size: 36 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 220 }, children: [t("生成日期：2026-07-01；本地语料最新抓取至：2026-06-30 23:40；狼大最新至：2026-06-30 22:33", { size: 20, color: "666666" })] }),
  p("说明：本文仅做论坛发言学习与交易复盘框架整理，不构成任何投资建议。个股部分只作为观察池和验证清单，不输出买卖指令。"),
  p("数据边界：本地 selected_users_posts.json 未检索到 2026-07-01 新发言；因此本文按当前能抓取到的最新本地增量，即 2026-06-30 晚间楼层整理。"),
  p("数据源规则：优先使用东方财富；东方财富接口断连、限流或返回空数据时，行情切换至腾讯财经。资金流若不可得，不用量价替代资金流结论，只保留为待验证项。"),

  h1("零、本次更新说明"),
  table(["项目", "结果", "处理"], dataRows, [1700, 5200, 2460]),

  h1("一、最新狼大发言要点"),
  table(["楼层", "主题", "原意整理", "学习要点"], wolfRows, [900, 1700, 4200, 2560]),

  h1("二、今天你问到的问题与分析结论"),
  table(["主题", "你的问题", "分析过程", "结果"], questionRows, [1450, 2350, 2800, 2760]),

  h1("三、新能源方向：聚和模板后的观察池"),
  p("聚和的启发不是“新能源标签”，而是技术/产能落地和资金重新定价。狼大对电池的触发点是亿纬锂能业绩确认储能加速，叠加AIDC需要储能扩充。"),
  table(["标的", "方向", "量价/位置", "后续验证"], newEnergyRows, [1700, 1700, 2700, 3260]),
  bullet("降权：天赐材料可观察但狼大明确说不在；纯锂矿变量多，需看宁德小作文、非洲矿、澳洲出口等。"),
  bullet("风险：新能源不是无条件主线，必须看业绩、毛利率、储能订单和板块核心是否继续带队。"),

  h1("四、半导体材料：石英锚点后的稀缺卡位"),
  p("狼大最新材料口径：上半年设备验证，下半年材料看扩产后的持续消耗。材料消耗按投片与扩产放大，不是按最终good die简单推导。当前位置已经不低，很多票需要从“低位看逻辑”切换到“高位看量价”。"),
  table(["标的", "稀缺卡位", "当前状态", "后续验证"], materialRows, [1700, 2100, 2600, 2960]),
  bullet("核心但等回踩：鼎龙股份、安集科技、雅克科技、南大光电、彤程新材、江丰电子。"),
  bullet("强度标杆但不宜只看逻辑：中船特气、华特气体、有研新材、神工股份、有研硅、中巨芯等已明显加速方向。"),

  h1("五、AIDC方向：疑似大资金拿筹码的筛选"),
  p("狼大最新口径不是泛IDC，而是华为国算链、国产卡放量、自建AIDC、IDC转AIDC和Token工厂。量能放大只能作为假设，不能直接证明机构；真正确认要看龙虎榜机构席位、基金持仓、调研纪要和订单公告。"),
  table(["标的", "方向", "量能证据", "判断"], aidcRows, [1700, 1900, 2900, 2860]),
  bullet("降权：中恒电气被狼大明确排除为国算；东方国信“算，但我不做”，只能当情绪参考。"),
  bullet("降权：润泽、光环、奥飞、数据港等传统IDC辨识度有，但本轮量能与华为卡/Token工厂口径不够贴合。"),
  bullet("排序口径：科华数据 > 英维克 > 华丰科技 > 同飞股份 > 申菱环境 > 中科曙光 > 浪潮信息 > 瑞可达。最新修正：因资金流数据不可得，这个排序只代表量价和逻辑贴合度，不代表机构确认买入。"),

  h1("六、后续复盘规则"),
  h2("1. 必须补的外部证据"),
  bullet("公司公告：订单、客户、产能、认证、并购/技术转移、业绩预告。"),
  bullet("研报/调研纪要：收入占比、毛利率、客户导入、下游需求、量产节奏。"),
  bullet("行情数据：5/20/60/144日线、5日/20日成交量、10日/60日成交量、放量后是否缩量承接。"),
  h2("2. 狼大视角下的动作降级条件"),
  bullet("只有小作文、没有公告或业绩支撑：降为观察。"),
  bullet("已经明显加速且放量冲高：不再用低位逻辑解释。"),
  bullet("核心票走弱、板块无联动、跌破关键均线：先看防守和验证，不强行脑补。"),
  h2("3. 明日/下次更新重点"),
  bullet("如果抓到 2026-07-01 新发言，优先更新AIDC、Token工厂、华为国算链和材料/新能源之间的资金切换。"),
  bullet("对AIDC观察池补龙虎榜与机构资金；对半导体材料补公告和中报预告；对新能源补储能订单和毛利率。"),
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Microsoft YaHei", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "Microsoft YaHei" }, paragraph: { spacing: { before: 260, after: 180 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 27, bold: true, font: "Microsoft YaHei" }, paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 220 } } } }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 },
      },
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [t("第 ", { size: 18 }), new TextRun({ children: [PageNumber.CURRENT], font: "Microsoft YaHei", size: 18 }), t(" 页", { size: 18 })] })] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outFile, buffer);
  console.log(outFile);
});
