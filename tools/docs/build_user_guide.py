"""Build the Chinese operator guide with the product's bundled Chinese font."""
from pathlib import Path
from xml.sax.saxutils import escape
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / 'output/pdf/凯乐提交文件工具_用户操作指南_中文版.pdf'
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
FONT = 'CFL-CN'
pdfmetrics.registerFont(TTFont(FONT, str(ROOT / 'ZTools.Offline/Fonts/NotoSansSC-Regular.ttf')))
pdfmetrics.registerFontFamily(FONT, normal=FONT, bold=FONT, italic=FONT, boldItalic=FONT)
BLUE = colors.HexColor('#0866AC')
NAVY = colors.HexColor('#173E5E')
INK = colors.HexColor('#243E52')
MUTED = colors.HexColor('#62798B')
PALE = colors.HexColor('#EDF5FB')
LINE = colors.HexColor('#D7E3ED')
WHITE = colors.white
PAGE_W, PAGE_H = A4
WIDTH = PAGE_W - 84
LOGO = str(ROOT / 'ZTools.Standalone/wwwroot/cfl-logo.jpg')

styles = {
    'body': ParagraphStyle('body', fontName=FONT, fontSize=10, leading=16, textColor=INK,
                           wordWrap='CJK', spaceAfter=7, splitLongWords=True),
    'small': ParagraphStyle('small', fontName=FONT, fontSize=8.8, leading=14, textColor=MUTED,
                            wordWrap='CJK', spaceAfter=5),
    'cell': ParagraphStyle('cell', fontName=FONT, fontSize=9.5, leading=15, textColor=INK, wordWrap='CJK'),
    'th': ParagraphStyle('th', fontName=FONT, fontSize=9.5, leading=15, textColor=WHITE, wordWrap='CJK'),
    'h1': ParagraphStyle('h1', fontName=FONT, fontSize=21, leading=30, textColor=NAVY, spaceAfter=12),
    'h2': ParagraphStyle('h2', fontName=FONT, fontSize=12.8, leading=20, textColor=BLUE, spaceBefore=10, spaceAfter=7),
    'cover': ParagraphStyle('cover', fontName=FONT, fontSize=30, leading=43, textColor=NAVY, spaceAfter=4),
    'subtitle': ParagraphStyle('subtitle', fontName=FONT, fontSize=13, leading=22, textColor=BLUE, spaceAfter=12),
}

def p(text, style='body'):
    return Paragraph(text, styles[style])

def heading(number, title):
    item = p(f'<font color="#0866AC">{number}</font>  {title}', 'h1')
    item.guide_bookmark = ('section' + number, title)
    return item

def table(headers, rows, widths, size=9.5):
    data = [[p(escape(x), 'th') for x in headers]]
    data += [[p(str(x), 'cell') for x in row] for row in rows]
    t = Table(data, colWidths=widths, repeatRows=1, hAlign='LEFT')
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 10), ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8), ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, colors.HexColor('#F5F8FB')]),
        ('LINEBELOW', (0, 0), (-1, 0), 0.6, BLUE),
        ('LINEBELOW', (0, 1), (-1, -1), 0.4, LINE),
    ]))
    return t

def note(title, text):
    box = Table([[p(title, 'h2')], [p(text)]], colWidths=[WIDTH], hAlign='LEFT')
    box.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), PALE), ('BOX', (0, 0), (-1, -1), 0.5, LINE),
        ('LEFTPADDING', (0, 0), (-1, -1), 13), ('RIGHTPADDING', (0, 0), (-1, -1), 13),
        ('TOPPADDING', (0, 0), (-1, 0), 8), ('BOTTOMPADDING', (0, 0), (-1, 0), 0),
        ('TOPPADDING', (0, 1), (-1, 1), 0), ('BOTTOMPADDING', (0, 1), (-1, 1), 8),
    ]))
    return box

def step(number, title, text):
    return p(f'<font color="#0866AC">{number:02d}  {title}</font><br/>{text}')

class InterfaceMap(Flowable):
    def __init__(self):
        Flowable.__init__(self)
        self.width, self.height = WIDTH, 225

    def draw(self):
        c = self.canv
        def text(x, y, value, size=9, color=INK):
            c.setFillColor(color); c.setFont(FONT, size); c.drawString(x, y, value)
        def rect(x, y, w, h, fill=WHITE):
            c.setFillColor(fill); c.setStrokeColor(LINE); c.roundRect(x, y, w, h, 5, fill=1, stroke=1)
        rect(0, 0, WIDTH, 220)
        c.setFillColor(NAVY); c.roundRect(0, 192, WIDTH, 28, 5, fill=1, stroke=0)
        text(12, 201, 'CFL  凯乐提交文件工具', 10, WHITE)
        text(WIDTH-127, 201, '刷新项目    退出程序', 9, WHITE)
        left, center, right = 125, 230, WIDTH-375
        rect(7, 8, left-6, 176, PALE)
        regions = [
            ('01 项目管理', ['新建项目 / 打开项目', '在 Excel 中输入数据']),
            ('02 数据计算', ['计算项目数据', '计算公用工程用量']),
            ('03 参数与选择', ['设备 / 电机 / 管仪', '剖面 / 公用工程 / 项目']),
            ('04 图纸输出', ['选择 GA / PID / SEC', '生成 DWG ＋ PDF']),
        ]
        for idx, (title, lines) in enumerate(regions):
            y = 169 - idx*42
            text(15, y, title, 9.2, BLUE)
            text(15, y-14, lines[0], 8.5)
            text(15, y-26, lines[1], 8.5)
        x = 130
        rect(x, 155, center, 29, PALE)
        text(x+8, 165, '总装外形图    管路仪表图    剖面图', 9)
        rect(x, 65, center, 84)
        text(x+65, 111, '图纸预览区域', 13, BLUE)
        text(x+50, 88, '滚轮缩放 · 拖动平移', 10, MUTED)
        rect(x, 30, center, 29, PALE); text(x+10, 40, '已生成图纸 / 保存位置 / 预览与下载', 9)
        text(x+9, 15, '操作记录：查看进度和错误提示', 8.5, MUTED)
        x = 368
        rect(x, 8, right, 176, PALE)
        text(x+10, 166, '设备清单', 10, BLUE)
        for idx in range(3):
            y = 133 - idx*30
            if idx == 0: rect(x+6, y-8, right-12, 27, WHITE)
            c.setStrokeColor(BLUE); c.rect(x+13, y, 8, 8, fill=0, stroke=1)
            if idx < 2:
                c.line(x+14, y+4, x+16, y+2); c.line(x+16, y+2, x+20, y+7)
            text(x+28, y, f'设备 {idx+1:02d}', 9)
        text(x+10, 47, '勾选：选择出图设备', 8.5)
        text(x+10, 31, '点行：选择修改对象', 8.5)
        text(x+10, 15, '内部计算结果：只读', 8.5)

class CheckItem(Flowable):
    def __init__(self, text):
        Flowable.__init__(self); self.paragraph=p(text); self.width=WIDTH
        _, h=self.paragraph.wrap(WIDTH-25, 100); self.height=h+8
    def draw(self):
        self.canv.setStrokeColor(BLUE); self.canv.setLineWidth(.8)
        self.canv.rect(1, self.height-16, 9, 9)
        self.paragraph.drawOn(self.canv, 23, 5)

class GuideDoc(SimpleDocTemplate):
    def afterFlowable(self, flowable):
        if hasattr(flowable, 'guide_bookmark'):
            key, title=flowable.guide_bookmark
            self.canv.bookmarkPage(key); self.canv.addOutlineEntry(title, key, 0)

def page_furniture(c, doc):
    c.saveState()
    if doc.page > 1:
        c.drawImage(LOGO, 42, PAGE_H-57, width=31, height=31, preserveAspectRatio=True, mask='auto')
        c.setFont(FONT, 9); c.setFillColor(MUTED)
        c.drawString(83, PAGE_H-39, '凯乐提交文件工具  /  用户操作指南')
        c.setStrokeColor(LINE); c.line(42, PAGE_H-68, PAGE_W-42, PAGE_H-68)
    c.setStrokeColor(LINE); c.line(42, 43, PAGE_W-42, 43)
    c.setFont(FONT, 8); c.setFillColor(MUTED)
    c.drawString(42, 28, 'CFL 凯乐 · 中文版 · 2026年10月27日')
    c.drawRightString(PAGE_W-42, 28, f'{doc.page:02d} / 08')
    c.restoreState()

story=[]
add=story.extend

# 1: Cover and quick route.
from reportlab.platypus import Image
logo=Image(LOGO, width=67, height=67)
brand=Table([[logo,p('CFL 凯乐<br/><font color="#62798B" size="10">提交文件工具 · 操作说明</font>', 'subtitle')]],colWidths=[82,WIDTH-82])
brand.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),0)]))
add([brand,Spacer(1,16),p('凯乐提交文件工具','cover'),p('用户操作指南','cover'),Spacer(1,8),
     p('Windows 独立版 · 附手机离线版操作说明','subtitle'),
     p('文档版本 1.0　　更新日期：2026年10月27日','small'),Spacer(1,14),p('从这六步开始','h2')])
cards=[]
for number,title,body in [
    ('01','新建或打开项目','选择项目位置，生成或打开 .xlsx 项目。'),
    ('02','填写设备输入','编辑 input，保存后返回工作台。'),
    ('03','复核内部参数','先选设备，再修改管仪、剖面和公用工程参数。'),
    ('04','计算项目数据','等待“计算完成，可以生成图纸”。'),
    ('05','选择设备与图纸','勾选设备，再选择 GA、PID 或 SEC。'),
    ('06','生成并核对结果','生成 DWG＋PDF，检查材料清单与保存位置。'),
]:
    cards.append(p(f'<font color="#0866AC" size="18">{number}</font>  {title}<br/><font size="9">{body}</font>'))
grid=Table([cards[:3],cards[3:]],colWidths=[WIDTH/3]*3)
grid.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),PALE),('BOX',(0,0),(-1,-1),.5,LINE),('INNERGRID',(0,0),(-1,-1),.5,WHITE),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),12),('RIGHTPADDING',(0,0),(-1,-1),12),('TOPPADDING',(0,0),(-1,-1),11),('BOTTOMPADDING',(0,0),(-1,-1),9)]))
add([grid,Spacer(1,15),p('查阅目录','h2')])
toc=[('02','安装与项目创建'),('03','主界面与参数入口'),('04','Excel 输入与保存'),('05','SEC 剖面与材料规则'),('06','计算、出图与归档'),('07','常见问题与处理'),('08','手机离线版与交付核对')]
tocrows=[]
for i in range(0,len(toc),2):
    row=[]
    for n,title in toc[i:i+2]: row.append(p(f'<link href="#section{n}" color="#0866AC">{n}　{title}</link>'))
    if len(row)==1: row.append('')
    tocrows.append(row)
contents=Table(tocrows,colWidths=[WIDTH/2]*2);contents.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),0),('VALIGN',(0,0),(-1,-1),'TOP')]))
add([contents,Spacer(1,7),p('首次使用建议先阅读第 2、4 页；轴承架及关联零件材质规则见第 5 页。','small'),PageBreak()])

# 2: Installation and project lifecycle.
add([heading('02','安装与项目创建'),table(['项目','使用要求'],[
    ('运行环境','Windows 10（1809 或更新版本）或 Windows 11，x64 环境。'),
    ('桌面 Excel','完整使用 Windows 版的参数读取、Excel 编辑和出图功能，需要安装可正常使用的桌面版 Microsoft Excel。'),
    ('浏览器与图形','工作台通过 Windows 默认浏览器打开；浏览器须支持 WebGL 图纸预览。'),
    ('随包组件','安装包已含 .NET 运行库、数据库和 DWG 模板。无需安装 AutoCAD，也无需启用 Excel 宏。'),
],[100,WIDTH-100]),p('安装与启动','h2'),
    step(1,'运行安装程序','双击交付的 Windows Setup 安装程序，按中文向导完成安装；从桌面或开始菜单打开“凯乐提交工具”。'),
    step(2,'新建项目','点击“新建项目”，选择自己的文档或项目文件夹。程序创建含有 input 工作表的 ZTools.xlsx 项目副本。'),
    step(3,'继续已有项目','点击“打开项目”，选择已保存的 .xlsx。启动程序本身不会自动新建项目或打开 Excel。'),
    note('旧版 .xls 的处理','打开旧版 .xls 项目时，程序会创建独立的 .xlsx 项目副本，保留原文件。请以工作台显示的项目路径为准，后续继续使用新副本。'),
    p('保存位置与升级','h2'),p('把项目放在可写的项目目录，不要放在程序安装目录，也不要直接编辑原始参考模板。升级前先保存 Excel、返回工作台，并点击“退出程序”。关闭浏览器页面不会自动停止程序。'),
    p('新项目自带输入模板；数据库、计算表及 GA/PID/SEC 模板由应用管理，无需另外选择。','small'),PageBreak()])

# 3: UI map, no stale screenshot.
add([heading('03','主界面与参数入口'),p('左侧负责项目、计算和出图；中间预览图纸；右侧显示设备清单和当前设备。'),InterfaceMap(),
    p('主界面区域示意（非截图）。按钮文字以当前程序界面为准。','small'),
    table(['参数入口','用途'],[
        ('设备参数','位号、泵版本、型号、工况、密封方案、轴承形式等。'),
        ('电机与底座','电机型号、厂家、基座号、电压和功率。'),
        ('管仪参数','修改当前设备的内部 PID 参数及有关方案。'),
        ('剖面材料','材质等级、逐件材料、轴承形式及轴承架关联材料。'),
        ('公用工程参数','修改当前设备的公用工程输入；灰色公式项只读。'),
        ('项目信息','客户、项目、订货号、编制/审核及有关计算设置。'),
        ('查看内部计算结果','查看 SEC 图块、材料清单和材质等级等结果；只读。'),
    ],[112,WIDTH-112]),Spacer(1,9),p('先单击右侧设备行，再打开参数窗口；修改后点击“保存到 Excel”。勾选框只决定本次出图范围，不会把参数修改批量应用到所有设备。'),PageBreak()])

# 4: XLSX input.
add([heading('04','Excel 输入与保存'),p('点击“在 Excel 中输入数据”打开项目。外部工作簿只有 input；内部 Sec-List、PID、Utility 通过应用的参数窗口修改，参数会随项目文件保存。'),
    table(['位置 / 字段','填写方法'],[
        ('第 21-50 行','每行一台设备，最多 30 台。泵型号从第 21 行开始连续填写，中间不得留空行。'),
        ('D 列 / 位号','填写设备位号；生成后按位号核对图纸。'),
        ('E 列 / 泵版本号','按实际设计版本选择，如 7版、8版或 10版；不要为了通过计算随意更改版本。'),
        ('N 列 / 泵型号','填写准确的泵型号，需有对应数据库记录。当前数据库支持 CZ、ZA、ZE、ZF 系列。'),
        ('O 列 / 材质','选择材料等级，作为默认材料输入；SEC 的具体选择还需复核“剖面材料”。'),
        ('U 列 / 基座号','按实际电机/底座配置选择；必须能与型号和电压匹配。'),
        ('V 列 / 电机电压','必须包含单位 V，例如 380V。'),
        ('AA 列 / 轴承形式','选择正确的轴承形式，并在剖面参数中核对。'),
        ('I53 / 法兰标准','从模板下拉列表选择受支持的法兰标准；有关设置也可在“项目信息”中修改。'),
    ],[116,WIDTH-116]),
    p('完成编辑后的操作','h2'),p('先在 Excel 中保存，然后点击工作台“保存并返回工作台”。也可保存后正常关闭 Excel，程序会在返回工作台时重新读取最后保存的内容。'),
    note('保持模板结构','保留 input 工作表名称、表头、公式和下拉区域。仅填写实际输入值。不要把数据库工作表补回项目，也不要用原始参考文件替代项目副本。'),Spacer(1,9),
    p('如果从其他 Excel 窗口修改了项目，先保存并关闭该副本，再用“刷新项目”或重新打开项目，随后重新计算。未保存的编辑无法从项目文件恢复。','small'),PageBreak()])

# 5: SEC business rule.
add([heading('05','SEC 剖面与材料规则'),p('先选中需要修改的设备，再点击“剖面材料”。这里的输入决定该设备的 SEC 剖面选择和材料清单。'),
    table(['输入项','对 SEC 的影响'],[
        ('材质等级（F）','决定材料数据库中的等级列；可覆盖 input 的默认材料等级。'),
        ('逐件材料（G-M）','分别覆盖轴、键、叶轮、叶轮口环、泵体口环、叶轮螺母和隔套等对应零件的材料。'),
        ('轴承形式（N）','参与 SEC 图块及材料清单的选择；按实际结构核对。'),
        ('轴承架材质（O）','指定时，同步应用到所选清单中出现的 330.01、330.02、330.03、360.01、360.02。'),
    ],[126,WIDTH-126]),Spacer(1,13),
    note('轴承架及关联零件共用指定材质','330.01 / 330.02 / 330.03<br/>360.01 / 360.02<br/><br/>上述零件只要出现在当前材料清单中，就使用同一个指定的轴承架材质；并不表示每种泵的清单都包含全部五个零件。'),
    p('指定与留空的区别','h2'),
    table(['场景','操作','结果'],[
        ('指定材质','轴承架材质填写 HT250','清单中上述五类零件均使用 HT250。'),
        ('恢复数据库材料','把轴承架材质留空并保存','各零件分别使用所选材料等级的数据库材料，不强制统一。'),
    ],[98,142,WIDTH-240]),
    p('修改后的复核顺序','h2'),p('保存参数 → 点击“计算项目数据” → 查看当前设备的 SEC 图块和材料清单 → 重新生成 SEC → 核对 DWG/PDF 的材料行。旧图纸不会因参数修改而自动重写。'),
    p('SEC 图块名称与材料清单名称可能不同，例如图块带 WO 等后缀；以程序计算结果为准。内部计算结果为只读，请通过参数入口修改输入。','small'),PageBreak()])

# 6: Calculation and delivery.
add([heading('06','计算、出图与归档'),
    step(1,'保存全部修改','完成 Excel 输入并返回工作台，保存设备、管仪、剖面和公用工程参数。'),
    step(2,'计算项目数据','点击“计算项目数据”，等待完成提示。计算失败或取消后不能生成新图纸；修正后重新计算。需要时另点“计算公用工程用量”。'),
    step(3,'选择本次出图范围','在右侧勾选设备；在左侧勾选 GA 总装外形图、PID 管路仪表图、SEC 剖面图及明细表。'),
    step(4,'生成并确认保存','点击“生成 DWG ＋ PDF”。成功后显示“图纸生成完成”和保存位置，点击“确定”返回工作台。'),
    step(5,'预览与核对','从“已生成图纸”列表点击“预览”。滚轮缩放、拖动平移，“适合窗口”恢复完整视图。DWG/PDF 链接可另下载副本。'),
    table(['文件','用途'],[
        ('DWG','保留可编辑图纸，用于后续工程复核。'),
        ('PDF','用于阅读、审阅和打印，当前输出为 A3 横向矢量图纸。'),
        ('项目 .xlsx','保存输入与内部参数；继续修改或备份项目时必须保留。'),
    ],[95,WIDTH-95]),Spacer(1,10),
    note('保存位置','结果保存在项目目录内：<br/>Output-Standalone / 生成批次 / GA、PID、SEC<br/>以“图纸生成完成”对话框及输出区域显示的位置为准。每次生成创建新的批次目录。'),
    p('出图与打印设置','h2'),p('可设置窗口坐标、页边距和份数。默认窗口为 (0,0)-(420,297)，无特殊要求保留默认。份数控制 PDF 中的重复页数；程序不会直接向实体打印机发送任务。打印请在 PDF 阅读器中操作。'),
    p('顶部图纸标签用于查看模板；正式结果请从“已生成图纸”列表打开。交付前确认没有空白模板、裁切文字或错误设备位号。','small'),PageBreak()])

# 7: Practical troubleshooting.
add([heading('07','常见问题与处理'),table(['现象','处理方法'],[
    ('生成按钮不可用','先保存 Excel 并返回工作台，再运行“计算项目数据”。计算成功后选择设备和图纸类型。参数有变化时需要重新计算。'),
    ('项目被占用 / 无法保存','关闭其他程序中打开的同一项目副本；确认项目目录可写。若程序保留着 Excel 编辑窗口，先保存并返回。'),
    ('泵型号 / 基座记录缺失','按错误提示核对型号、版本、基座号和电压。数据库缺失时请由项目维护人员核查，不要为了通过计算改用其他版本。'),
    ('ZE 7版 查询失败','ZE 的 7版使用 ZE7 数据库。缺少对应型号/基座记录时会停止，不会自动改查其他 ZE 版本。'),
    ('SEC 材料清单找不到','核对“剖面材料”中的等级和轴承形式。缺失或重复的清单、零件数据需要维护人员核查，不能随意猜选。'),
    ('Excel 关闭后提示恢复','工作台重新读取最后保存的内容。确认保存值正确后重新计算；未保存的修改不能恢复。'),
    ('计算长时间未完成','可点击“取消计算”。程序也会对长时间无响应的计算停止处理。取消或失败时，项目原文件保持不变。'),
    ('浏览器 / 预览异常','先确认程序仍在运行；退出后从快捷方式重新打开。浏览器须支持 WebGL。若 PDF 提示不支持的图元，保留错误提示供维护人员检查模板。'),
],[125,WIDTH-125]),
    p('反馈问题时准备这些信息','h2'),p('复制完整错误文字，记录设备位号、泵型号、泵版本、基座号、电压和出错步骤。保留问题项目副本，避免在原文件上反复尝试不确定的改动。'),
    note('修正以后再出图','修正输入或数据库问题后，重新计算并重新生成图纸。不要把失败前的旧批次当作本次参数对应的正式结果。'),PageBreak()])

# 8: Platform availability and checklist.
add([heading('08','手机离线版与交付核对'),table(['平台','安装与使用'],[
    ('Windows','本手册第 2-7 页为主要操作流程。需要桌面版 Excel 完成 Windows 参数读取、输入编辑和出图。'),
    ('Android','提供 ARM64 APK，适用 Android 8 或更新版本。将交付的 APK 复制到手机，通过文件应用打开并按系统提示安装。'),
    ('iPhone','当前提供 Mac 构建套件，不是已签名可安装的 IPA。需由具备 Mac 和 Apple 签名条件的人员生成并分发 IPA；安装方式按实际签名方案执行。'),
],[88,WIDTH-88]),
    p('手机中的操作顺序','h2'),p('手机应用为本机离线版，不连接 Windows 服务，也不依赖桌面 Excel。新建或导入项目，选择当前设备，编辑参数，计算成功后生成 DWG/PDF；通过“导出”保存到手机文件应用。新项目与 Excel 导出使用 .xlsx，旧 .xls 项目可导入。'),
    p('交付前逐项核对','h2')])
for text in [
    '设备位号、设备数量和本次勾选范围正确。',
    '泵版本、型号、基座号、电压与实际设计一致。',
    'SEC 材料等级、逐件覆盖值和轴承形式已复核。',
    '指定轴承架材质时，330.01-330.03 与 360.01/360.02 的材料行已核对。',
    '已保存项目 .xlsx，且最后一次计算成功。',
    'DWG/PDF 对应本批次，图框、文字、明细表完整，设备信息一致。',
    '已记录保存位置，并备份项目文件及所需输出文件。',
]: story.append(CheckItem(text))
add([Spacer(1,9),note('完成后','点击“退出程序”。备份建议保留整个项目文件夹，至少包含项目 .xlsx 和需要留存的输出批次。以后修改时，从应用重新打开项目并重新计算。')])

doc=GuideDoc(str(OUTPUT),pagesize=A4,rightMargin=36,leftMargin=36,topMargin=83,bottomMargin=57,
             title='凯乐提交文件工具 用户操作指南（中文版）',author='CFL 凯乐',subject='Windows 独立版操作、XLSX 项目与 SEC 材料规则',pageCompression=1)
doc.build(story,onFirstPage=page_furniture,onLaterPages=page_furniture)
from pypdf import PdfReader
reader=PdfReader(str(OUTPUT))
assert len(reader.pages)==8, f'Expected 8 pages, got {len(reader.pages)}'
all_text='\n'.join(page.extract_text() or '' for page in reader.pages)
for text in ['用户操作指南','360.01','360.02','330.03','新建项目','保存并返回工作台','ZE7','iPhone']:
    assert text in all_text, text
assert '\ufffd' not in all_text
print(f'Created {OUTPUT} ({len(reader.pages)} pages, {OUTPUT.stat().st_size:,} bytes)')
