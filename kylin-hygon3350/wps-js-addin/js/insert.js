/**
 * 公文版头、版记、红线、日期、符号与电子公章插入模块 (移植自 Ribbon1.Insert.cs + 公文部件.dotx)
 * 在银河麒麟 Linux 下使用原生 JSAPI 直接构建版头/版记/装订线，不依赖 Windows COM BuildingBlocks
 */
var GwInsert = {
    CIRCLED_NUMS: [
        "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩",
        "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳"
    ],

    /**
     * 插入版记表格 (1=单行版记[印发机关+日期], 2=双行版记[抄送机关 + 印发机关+日期])
     * 严格还原原项目 InsertLine(int LineNumber) 的尺寸与边框规范
     */
    insertBanjiTable: function (lineCount) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var rows = (!lineCount || lineCount < 1) ? 1 : lineCount;
        var doc = app.ActiveDocument;
        var cfg = GwConfig.load();

        try {
            var tbl = doc.Tables.Add(app.Selection.Range, rows, 1);
            try {
                tbl.Rows.RelativeVerticalPosition = WpsEnum.wdRelativeVerticalPositionMargin;
                tbl.Rows.VerticalPosition = WpsEnum.wdFrameBottom;
            } catch (ePos) {}
            tbl.PreferredWidth = 443;
            tbl.Rows.Height = 28.5;
            tbl.Rows.Alignment = WpsEnum.wdAlignRowLeft;

            // 上下粗线 1.00pt，内部横线 0.75pt
            tbl.Borders.Item(WpsEnum.wdBorderTop).LineStyle = WpsEnum.wdLineStyleSingle;
            tbl.Borders.Item(WpsEnum.wdBorderTop).LineWidth = WpsEnum.wdLineWidth100pt;
            tbl.Borders.Item(WpsEnum.wdBorderBottom).LineStyle = WpsEnum.wdLineStyleSingle;
            tbl.Borders.Item(WpsEnum.wdBorderBottom).LineWidth = WpsEnum.wdLineWidth100pt;
            tbl.Borders.Item(WpsEnum.wdBorderLeft).LineStyle = WpsEnum.wdLineStyleNone;
            tbl.Borders.Item(WpsEnum.wdBorderRight).LineStyle = WpsEnum.wdLineStyleNone;

            if (rows > 1) {
                tbl.Borders.Item(WpsEnum.wdBorderHorizontal).LineStyle = WpsEnum.wdLineStyleSingle;
                tbl.Borders.Item(WpsEnum.wdBorderHorizontal).LineWidth = WpsEnum.wdLineWidth075pt;
                tbl.Cell(1, 1).Range.Text = "  抄送：○○○○○○。";
                tbl.Cell(2, 1).Range.Text = "  ○○○○○○办公室                        " + this.getTodayCnString() + "印发";
            } else {
                tbl.Cell(1, 1).Range.Text = "  ○○○○○○办公室                        " + this.getTodayCnString() + "印发";
            }

            tbl.Range.Cells.VerticalAlignment = WpsEnum.wdCellAlignVerticalCenter;
            tbl.Range.Font.NameFarEast = cfg.fonts.banji.name;
            tbl.Range.Font.Name = cfg.fonts.banji.name;
            tbl.Range.Font.Size = parseFloat(cfg.fonts.banji.size);
            app.StatusBar = rows > 1 ? "已插入标准双行版记（含抄送机关与印发机关）" : "已插入标准单行版记";
        } catch (e) {
            alert("插入版记时出错: " + e.message);
        }
    },

    /**
     * 插入版头：单行发文机关红头 (还原 公文部件.dotx -> 版头:单行发文机关)
     */
    insertRedHeadSingle: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var sel = app.Selection;
            sel.HomeKey(WpsEnum.wdStory);
            sel.InsertParagraphBefore();
            sel.HomeKey(WpsEnum.wdStory);

            // 1. 发文机关标志 (红色方正小标宋简体 42pt 居中)
            sel.Font.NameFarEast = "方正小标宋简体";
            sel.Font.Name = "方正小标宋简体";
            sel.Font.Size = 42;
            sel.Font.Bold = 0;
            sel.Font.Color = WpsEnum.wdColorRed;
            sel.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
            sel.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            sel.ParagraphFormat.FirstLineIndent = 0;
            sel.ParagraphFormat.SpaceBefore = 24;
            sel.ParagraphFormat.SpaceAfter = 18;
            sel.TypeText("请在此输入发文单位名称文件");
            sel.TypeParagraph();

            // 2. 发文字号与红色反线 (三号仿宋居中 + 2.25磅红底线)
            sel.Font.NameFarEast = "仿宋";
            sel.Font.Name = "仿宋";
            sel.Font.Size = 16;
            sel.Font.Color = WpsEnum.wdColorBlack;
            sel.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
            sel.ParagraphFormat.SpaceBefore = 6;
            sel.ParagraphFormat.SpaceAfter = 0;
            var year = new Date().getFullYear();
            sel.TypeText("○○发〔" + year + "〕1号");

            var btm = sel.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom);
            btm.LineStyle = WpsEnum.wdLineStyleSingle;
            btm.LineWidth = WpsEnum.wdLineWidth225pt;
            btm.Color = WpsEnum.wdColorRed;

            sel.TypeParagraph();
            sel.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom).LineStyle = WpsEnum.wdLineStyleNone;
            app.StatusBar = "已插入公文版头：单行发文机关标志与发文字号红线";
        } catch (e) {
            alert("插入单行发文机关版头出错: " + e.message);
        }
    },

    /**
     * 插入版头：联合发文机关红头 (还原 公文部件.dotx -> 版头:联合发文机关 2列无边框表格)
     */
    insertRedHeadMulti: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var doc = app.ActiveDocument;
            var sel = app.Selection;
            sel.HomeKey(WpsEnum.wdStory);
            sel.InsertParagraphBefore();
            sel.HomeKey(WpsEnum.wdStory);

            var tbl = doc.Tables.Add(sel.Range, 1, 2);
            tbl.Borders.Enable = 0;
            tbl.Rows.Alignment = WpsEnum.wdAlignRowCenter;

            var c1 = tbl.Cell(1, 1);
            c1.Range.Text = "此处输入主办单位名称\r此处输入联合发文单位名称";
            c1.Range.Font.NameFarEast = "方正小标宋简体";
            c1.Range.Font.Size = 26;
            c1.Range.Font.Color = WpsEnum.wdColorRed;
            c1.Range.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphDistribute;
            c1.VerticalAlignment = WpsEnum.wdCellAlignVerticalCenter;

            var c2 = tbl.Cell(1, 2);
            c2.Range.Text = "文件";
            c2.Range.Font.NameFarEast = "方正小标宋简体";
            c2.Range.Font.Size = 44;
            c2.Range.Font.Color = WpsEnum.wdColorRed;
            c2.Range.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
            c2.VerticalAlignment = WpsEnum.wdCellAlignVerticalCenter;

            // 跳出表格插入发文字号及红色间隔线
            sel.MoveDown(WpsEnum.wdLine, 3);
            sel.TypeParagraph();
            sel.Font.NameFarEast = "仿宋";
            sel.Font.Size = 16;
            sel.Font.Color = WpsEnum.wdColorBlack;
            sel.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
            var year = new Date().getFullYear();
            sel.TypeText("○○发〔" + year + "〕1号");
            var btm = sel.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom);
            btm.LineStyle = WpsEnum.wdLineStyleSingle;
            btm.LineWidth = WpsEnum.wdLineWidth225pt;
            btm.Color = WpsEnum.wdColorRed;
            sel.TypeParagraph();
            sel.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom).LineStyle = WpsEnum.wdLineStyleNone;

            app.StatusBar = "已插入公文版头：联合发文机关";
        } catch (e) {
            alert("插入联合发文机关版头出错: " + e.message);
        }
    },

    /**
     * 插入红色间隔线（反线 2.25磅）
     */
    insertRedDividerLine: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            var btm = app.Selection.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom);
            btm.LineStyle = WpsEnum.wdLineStyleSingle;
            btm.LineWidth = WpsEnum.wdLineWidth225pt;
            btm.Color = WpsEnum.wdColorRed;
            app.StatusBar = "已插入红色间隔线 (2.25磅)";
        } catch (e) {
            alert("插入红色间隔线出错: " + e.message);
        }
    },

    /**
     * 插入信函格式上沿双红线 / 下沿单红线
     */
    insertLetterLine: function (isTop) {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            var bdr = app.Selection.ParagraphFormat.Borders.Item(isTop ? WpsEnum.wdBorderTop : WpsEnum.wdBorderBottom);
            bdr.LineStyle = isTop ? WpsEnum.wdLineStyleDouble : WpsEnum.wdLineStyleSingle;
            bdr.LineWidth = isTop ? WpsEnum.wdLineWidth150pt : WpsEnum.wdLineWidth100pt;
            bdr.Color = WpsEnum.wdColorRed;
            app.StatusBar = isTop ? "已插入信函版头上沿红线" : "已插入信函版头下沿红线";
        } catch (e) {
            alert("插入函线出错: " + e.message);
        }
    },

    /**
     * 插入左侧竖排装订线
     */
    insertBindingLine: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var hdr = app.ActiveDocument.Sections.Item(1).Headers.Item(WpsEnum.wdHeaderFooterPrimary);
            var shp = hdr.Shapes.AddTextbox(2, mmToPt(12), mmToPt(60), mmToPt(10), mmToPt(175));
            shp.Line.Visible = 0;
            shp.Fill.Visible = 0;
            shp.TextFrame.TextRange.Text = "┆┆┆┆┆┆┆┆┆┆┆┆装┆┆┆┆┆┆┆┆┆┆┆┆订┆┆┆┆┆┆┆┆┆┆┆┆线┆┆┆┆┆┆┆┆┆┆┆┆";
            shp.TextFrame.TextRange.Font.NameFarEast = "宋体";
            shp.TextFrame.TextRange.Font.Size = 10.5;
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
            app.StatusBar = "已在页面左侧插入公文竖排装订线";
        } catch (e) {
            alert("插入装订线出错: " + e.message);
        }
    },

    getTodayCnString: function () {
        var d = new Date();
        return d.getFullYear() + "年" + (d.getMonth() + 1) + "月" + d.getDate() + "日";
    },

    /**
     * 插入公文格式当前日期 (YYYY年M月D日)
     */
    insertCnDate: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var cfg = GwConfig.load();
        GwFonts.applyFontSpec(cfg.fonts.normal);
        try {
            app.Selection.TypeText(this.getTodayCnString());
            app.Selection.InsertParagraphAfter();
            app.Selection.Collapse(WpsEnum.wdCollapseEnd);
            app.StatusBar = "已插入公文格式日期";
        } catch (e) {
            alert("插入日期出错: " + e.message);
        }
    },

    /**
     * 在当前页之后插入一个独立的横向 A4 页面，并随后恢复纵向页面
     */
    insertLandscapePage: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            app.ActiveWindow.ActivePane.View.SeekView = WpsEnum.wdSeekMainDocument;
            app.Selection.InsertBreak(WpsEnum.wdSectionBreakNextPage);
            app.Selection.PageSetup.Orientation = WpsEnum.wdOrientLandscape;
            app.Selection.InsertBreak(WpsEnum.wdSectionBreakNextPage);
            app.Selection.PageSetup.Orientation = WpsEnum.wdOrientPortrait;
            app.StatusBar = "已插入独立横向A4页面分节";
        } catch (e) {
            alert("插入横页失败: " + e.message);
        }
    },

    /**
     * 插入特殊符号
     */
    insertSymbolText: function (sym) {
        var app = getWpsApp();
        if (!app || !app.Selection || !sym) return;
        try {
            app.Selection.TypeText(sym);
            app.StatusBar = "已插入符号: " + sym;
        } catch (e) {
            alert("插入符号失败: " + e.message);
        }
    },

    /**
     * 将 [2026] 或 ［2026］ 替换为 〔2026〕
     */
    replaceSquareBracketsToHex: function () {
        GwFormat.wordReplace(true, "\\[([0-9]{1,4})\\]", "〔\\1〕", false);
        GwFormat.wordReplace(true, "［([0-9]{1,4})］", "〔\\1〕", false);
        var app = getWpsApp();
        if (app) app.StatusBar = "已将年份方括号 [] 批量替换为六角括号 〔〕";
    },

    /**
     * 将 【2026】 替换为 〔2026〕
     */
    replaceBlackBracketsToHex: function () {
        GwFormat.wordReplace(true, "【([0-9]{1,4})】", "〔\\1〕", false);
        var app = getWpsApp();
        if (app) app.StatusBar = "已将年份黑括号 【】 批量替换为六角括号 〔〕";
    },

    /**
     * 将选中的 1~20 数字转换为带圈序号 ①~⑳
     */
    convertNumberToCircled: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var raw = (app.Selection.Text || "").trim();
        var num = parseInt(raw, 10);
        if (isNaN(num) || num < 1 || num > 20) {
            alert("请先选中一个 1 ~ 20 之间的阿拉伯数字！");
            return;
        }
        app.Selection.Text = this.CIRCLED_NUMS[num - 1];
        app.StatusBar = "已转换为带圈序号: " + this.CIRCLED_NUMS[num - 1];
    },

    /**
     * 插入或链接电子公章图片并设为浮于文字上方
     */
    insertOfficialStamp: function (linkToFile, everyPage) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;

        // 优先调用本地麒麟桥接服务打开原生文件选择器，若未启动则使用 prompt 输入路径
        GwConfig.callBridge("/api/pick-file", {
            title: everyPage ? "请选择每页批量加盖的透明公章图片" : "请选择要插入的透明电子公章图片 (PNG/GIF/TIFF)",
            filter: "*.png *.gif *.tiff"
        }, function (err, resp) {
            var filePath = (resp && resp.path) ? resp.path : "";
            if (!filePath) {
                filePath = prompt("请输入电子公章图片的完整路径（支持透明PNG）：", "/home/kylin/公章.png");
            }
            if (!filePath) return;

            try {
                if (everyPage) {
                    var secs = app.ActiveDocument.Sections;
                    for (var i = 1; i <= secs.Count; i++) {
                        var hdr = secs.Item(i).Headers.Item(WpsEnum.wdHeaderFooterPrimary);
                        var shp = hdr.Shapes.AddPicture(filePath, false, true);
                        shp.LockAspectRatio = -1;
                        shp.Width = cmToPt(4.0);
                        shp.Height = cmToPt(4.0);
                        shp.WrapFormat.Type = WpsEnum.wdWrapBehind;
                        shp.Left = 340;
                        shp.Top = 480;
                    }
                    app.StatusBar = "已在每页批量加盖电子公章";
                } else {
                    var pic = app.Selection.Range.InlineShapes.AddPicture(filePath, !!linkToFile, !linkToFile);
                    try {
                        var sPic = pic.ConvertToShape();
                        sPic.WrapFormat.Type = WpsEnum.wdWrapFront;
                        sPic.ZOrder(WpsEnum.msoBringInFrontOfText);
                    } catch (eShape) {}
                    app.StatusBar = linkToFile ? "已以安全链接方式插入电子公章" : "已嵌入电子公章并设为浮于文字上方";
                }
            } catch (ex) {
                alert("插入电子公章失败，请确认图片路径存在且可读：\n" + ex.message);
            }
        });
    }
};
