/**
 * 版面与页面设置模块 (移植自 Ribbon1.Layout.cs & frmPaperSetup.cs)
 */
var GwLayout = {
    /**
     * 打开公文页面设置对话框
     */
    openPaperSetupDialog: function () {
        var url = window.location.href.replace(/index\.html.*$/, "ui/paper_setup.html");
        if (typeof wps !== "undefined" && wps.ShowDialog) {
            wps.ShowDialog(url, "公文页面设置 (GB/T 9704—2012)", 480, 460, false);
        } else {
            window.open(url, "_blank", "width=480,height=460");
        }
    },

    /**
     * 应用 GB/T 9704 公文版面参数到当前活动文档
     */
    applyGwPageSetup: function (customCfg) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) {
            alert("当前没有打开的公文文档！");
            return false;
        }
        var doc = app.ActiveDocument;
        var cfg = customCfg || GwConfig.load();

        var topPt = mmToPt(cfg.topMarginMm);
        var bottomPt = mmToPt(cfg.bottomMarginMm);
        var leftPt = mmToPt(cfg.leftMarginMm);
        var rightPt = mmToPt(cfg.rightMarginMm);
        var footPt = mmToPt(cfg.footerDistanceMm);

        try {
            // 设置正文默认样式字体
            var normalStyle = doc.Styles.Item(WpsEnum.wdStyleNormal);
            if (normalStyle) {
                normalStyle.Font.NameFarEast = cfg.fonts.normal.name;
                if (cfg.letterFont) {
                    normalStyle.Font.NameAscii = cfg.letterFont;
                    normalStyle.Font.Name = cfg.letterFont;
                }
                if (cfg.useGwGrid) {
                    normalStyle.Font.Size = 16;
                }
            }
        } catch (e) {
            console.warn("设置默认样式字体提示:", e);
        }

        try {
            var secCount = doc.Sections.Count;
            for (var i = 1; i <= secCount; i++) {
                var sec = doc.Sections.Item(i);
                var ps = sec.PageSetup;
                ps.TopMargin = topPt;
                ps.BottomMargin = bottomPt;
                ps.LeftMargin = leftPt;
                ps.RightMargin = rightPt;
                ps.FooterDistance = footPt;

                if (ps.Orientation === WpsEnum.wdOrientPortrait) {
                    ps.PageWidth = cmToPt(21.0);
                    ps.PageHeight = cmToPt(29.7);
                    if (cfg.useGwGrid) {
                        ps.LayoutMode = WpsEnum.wdLayoutModeGrid;
                        ps.CharsLine = cfg.charsPerLine;
                        ps.LinesPage = cfg.linesPerPage;
                    }
                    if (cfg.useExactLineSpacing) {
                        sec.Range.ParagraphFormat.LineSpacingRule = WpsEnum.wdLineSpaceExactly;
                        sec.Range.ParagraphFormat.LineSpacing = parseFloat(cfg.lineSpacingPt);
                    }
                } else {
                    ps.PageWidth = cmToPt(29.7);
                    ps.PageHeight = cmToPt(21.0);
                }
            }

            if (typeof cfg.embedFonts === "boolean") {
                try {
                    doc.EmbedTrueTypeFonts = cfg.embedFonts;
                    doc.SaveSubsetFonts = cfg.embedFonts;
                } catch (e2) {}
            }
            if (typeof cfg.showHiddenText === "boolean") {
                try {
                    app.ActiveWindow.View.ShowHiddenText = cfg.showHiddenText;
                    app.Options.PrintHiddenText = false;
                } catch (e3) {}
            }
            app.StatusBar = "已应用国家标准公文页面版式 (上" + cfg.topMarginMm + "mm 下" + cfg.bottomMarginMm + "mm 左" + cfg.leftMarginMm + "mm 右" + cfg.rightMarginMm + "mm)";
            return true;
        } catch (err) {
            alert("应用公文页面设置时出错: " + err.message);
            return false;
        }
    },

    /**
     * 中文缩进：首行缩进 2 字符
     */
    indent2Chars: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            app.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
            app.StatusBar = "已设置首行缩进2字符";
        } catch (e) {
            alert("中文缩进出错: " + e.message);
        }
    },

    /**
     * 段落顶格：清除首行缩进、左缩进
     */
    paragraphFlushLeft: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            var pf = app.Selection.ParagraphFormat;
            pf.CharacterUnitFirstLineIndent = 0;
            pf.FirstLineIndent = 0;
            pf.LeftIndent = 0;
            pf.CharacterUnitLeftIndent = 0;
            app.StatusBar = "已设置段落顶格";
        } catch (e) {
            alert("段落顶格出错: " + e.message);
        }
    },

    /**
     * 删除页眉横线
     */
    removeHeaderLine: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            app.ScreenUpdating = false;
            var view = app.ActiveWindow.ActivePane.View;
            if (view.SeekView !== WpsEnum.wdSeekCurrentPageHeader) {
                view.SeekView = WpsEnum.wdSeekCurrentPageHeader;
            }
            app.Selection.WholeStory();
            app.Selection.Borders.Item(WpsEnum.wdBorderBottom).LineStyle = WpsEnum.wdLineStyleNone;
            view.SeekView = WpsEnum.wdSeekMainDocument;
            app.StatusBar = "已删除页眉横线";
        } catch (e) {
            try {
                var secs = app.ActiveDocument.Sections;
                for (var i = 1; i <= secs.Count; i++) {
                    secs.Item(i).Headers.Item(WpsEnum.wdHeaderFooterPrimary).Range.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom).LineStyle = WpsEnum.wdLineStyleNone;
                }
            } catch (e2) {
                alert("删除页眉横线失败: " + e.message);
            }
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 删除页眉内容与横线
     */
    deletePageHeader: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            app.ScreenUpdating = false;
            var secs = app.ActiveDocument.Sections;
            for (var i = 1; i <= secs.Count; i++) {
                var hdr = secs.Item(i).Headers.Item(WpsEnum.wdHeaderFooterPrimary).Range;
                hdr.Delete();
                hdr.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom).LineStyle = WpsEnum.wdLineStyleNone;
            }
            app.StatusBar = "已删除文档页眉及页眉横线";
        } catch (e) {
            alert("删除页眉失败: " + e.message);
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 编号转文字（全文自动编号转静态文本）
     */
    convertNumbersToText: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        if (!confirm("编号转换为文字后，现有的项目将转为静态文本，无法继续自动编号。\n确定要继续吗？")) {
            return;
        }
        try {
            app.ActiveDocument.Content.ListFormat.ConvertNumbersToText();
            app.StatusBar = "已将文档所有自动编号转换为静态文字";
        } catch (e) {
            alert("编号转文字出错: " + e.message);
        }
    }
};
