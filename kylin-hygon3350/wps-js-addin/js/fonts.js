/**
 * 公文字体与四级标题设置模块 (移植自 Ribbon1.Fonts.cs & frmFontChangeDialog.cs)
 */
var GwFonts = {
    CN_NUMS: ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
              "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十"],

    /**
     * 设置所选区域字体属性
     */
    applyFontSpec: function (fontSpec) {
        var app = getWpsApp();
        if (!app || !app.Selection || !fontSpec) return;
        try {
            var f = app.Selection.Font;
            f.NameFarEast = fontSpec.name;
            f.Name = fontSpec.name;
            f.Size = parseFloat(fontSpec.size);
            f.Bold = fontSpec.bold ? 1 : 0;
            f.Italic = fontSpec.italic ? 1 : 0;
        } catch (e) {
            console.error("applyFontSpec error:", e);
        }
    },

    /**
     * 公文大标题（二号方正小标宋简体，居中）
     */
    setMainTitle: function (customFontName) {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var cfg = GwConfig.load();
        var spec = JSON.parse(JSON.stringify(cfg.fonts.title));
        if (customFontName) {
            spec.name = customFontName;
            spec.size = 22.0;
        }
        this.applyFontSpec(spec);
        try {
            app.Selection.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
            app.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            app.Selection.ParagraphFormat.FirstLineIndent = 0;
            app.StatusBar = "已设为公文标题格式 (" + spec.name + " 二号居中)";
        } catch (e) {
            alert("设置公文标题出错: " + e.message);
        }
    },

    /**
     * 公文正文（三号仿宋，取消编号，大纲级别设为正文文本）
     */
    setNormalBody: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var cfg = GwConfig.load();
        this.applyFontSpec(cfg.fonts.normal);
        try {
            app.Selection.Range.ListFormat.RemoveNumbers(WpsEnum.wdNumberAllNumbers);
        } catch (e) {}
        try {
            app.Selection.ParagraphFormat.OutlineLevel = WpsEnum.wdOutlineLevelBodyText;
            app.Selection.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphJustify;
            app.StatusBar = "已设为公文正文格式 (" + cfg.fonts.normal.name + " 三号)";
        } catch (e2) {}
    },

    /**
     * 设置 1~4 级公文标题
     * level: 1 -> 一、 (三号黑体)
     * level: 2 -> （一） (三号楷体)
     * level: 3 -> 1． (三号仿宋)
     * level: 4 -> （1） (三号仿宋)
     */
    setHeadingLevel: function (level, startIndex) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument || !app.Selection) return;
        var cfg = GwConfig.load();
        var idx = parseInt(startIndex, 10) || 1;
        if (idx < 1) idx = 1;

        var fontMap = {
            1: cfg.fonts.title1,
            2: cfg.fonts.title2,
            3: cfg.fonts.title3,
            4: cfg.fonts.title4
        };
        this.applyFontSpec(fontMap[level] || cfg.fonts.title1);

        try {
            var tpl = app.ActiveDocument.ListTemplates.Add();
            var lvl = tpl.ListLevels.Item(1);
            lvl.TrailingCharacter = WpsEnum.wdTrailingNone;
            lvl.TextPosition = 0;
            lvl.NumberPosition = 0;
            lvl.StartAt = idx;

            if (level === 1) {
                lvl.NumberStyle = WpsEnum.wdListNumberStyleSimpChinNum3;
                lvl.NumberFormat = "%1、";
                app.Selection.ParagraphFormat.OutlineLevel = WpsEnum.wdOutlineLevel1;
            } else if (level === 2) {
                lvl.NumberStyle = WpsEnum.wdListNumberStyleSimpChinNum3;
                lvl.NumberFormat = "（%1）";
                app.Selection.ParagraphFormat.OutlineLevel = WpsEnum.wdOutlineLevel2;
            } else if (level === 3) {
                lvl.NumberFormat = "%1．";
                app.Selection.ParagraphFormat.OutlineLevel = WpsEnum.wdOutlineLevel3;
            } else if (level === 4) {
                lvl.NumberFormat = "（%1）";
                app.Selection.ParagraphFormat.OutlineLevel = WpsEnum.wdOutlineLevel4;
            }
            app.Selection.Range.ListFormat.ApplyListTemplate(tpl);
            app.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
            app.StatusBar = "已设置为公文 " + level + " 级标题";
        } catch (e) {
            // 若部分Linux WPS版本ListTemplate兼容受限，优雅回退为直接插入规范标题前缀
            var prefix = "";
            var cn = this.CN_NUMS[idx] || String(idx);
            if (level === 1) prefix = cn + "、";
            else if (level === 2) prefix = "（" + cn + "）";
            else if (level === 3) prefix = idx + "．";
            else prefix = "（" + idx + "）";
            app.Selection.InsertBefore(prefix);
            app.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
        }
    },

    /**
     * 快捷预设字体：版头(三号黑体)、签发人(三号楷体)、版记(四号仿宋)
     */
    setQuickRoleFont: function (role) {
        var cfg = GwConfig.load();
        if (role === "head") this.applyFontSpec(cfg.fonts.head);
        else if (role === "name") this.applyFontSpec(cfg.fonts.signerName);
        else if (role === "foot") this.applyFontSpec(cfg.fonts.banji);
    },

    /**
     * 切换着重号
     */
    toggleUnderDot: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            var f = app.Selection.Font;
            if (f.EmphasisMark !== WpsEnum.wdEmphasisMarkUnderSolidCircle) {
                f.EmphasisMark = WpsEnum.wdEmphasisMarkUnderSolidCircle;
                app.StatusBar = "已添加着重号";
            } else {
                f.EmphasisMark = WpsEnum.wdEmphasisMarkNone;
                app.StatusBar = "已取消着重号";
            }
        } catch (e) {
            alert("设置着重号出错: " + e.message);
        }
    },

    /**
     * 切换隐藏文字
     */
    toggleHiddenFont: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            var f = app.Selection.Font;
            f.Hidden = f.Hidden ? 0 : -1;
            app.StatusBar = f.Hidden ? "所选文字已设为隐藏" : "已取消文字隐藏";
        } catch (e) {
            alert("设置隐藏文字出错: " + e.message);
        }
    },

    /**
     * 全文替换中文字体 (例如 仿宋_GB2312 <=> 仿宋)
     */
    switchFont: function (oldFont, newFont) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var find = app.ActiveDocument.Content.Find;
            find.ClearFormatting();
            find.Font.NameFarEast = oldFont;
            find.Replacement.ClearFormatting();
            find.Replacement.Font.NameFarEast = newFont;
            find.Execute("", false, false, false, false, false, true,
                WpsEnum.wdFindContinue, true, "", WpsEnum.wdReplaceAll);
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            app.StatusBar = "字体【" + oldFont + "】已批量替换为【" + newFont + "】";
        } catch (e) {
            alert("替换字体时出错: " + e.message);
        }
    },

    /**
     * 更改全文数字(^#)或英文字母(^$)的西文字体
     */
    changeNumOrLetterFont: function (pattern, fontName) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var find = app.ActiveDocument.Content.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = pattern;
            find.Forward = true;
            find.Wrap = WpsEnum.wdFindContinue;
            find.Format = true;
            find.MatchByte = true;
            find.MatchWildcards = false;
            find.Replacement.Text = "^&";
            find.Replacement.Font.NameAscii = fontName;
            find.Replacement.Font.Name = fontName;
            find.Execute(undefined, undefined, undefined, undefined, undefined,
                undefined, undefined, undefined, undefined, undefined, WpsEnum.wdReplaceAll);
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
        } catch (e) {
            alert("更改数字/字母字体出错: " + e.message);
        }
    },

    /**
     * 打开字体替换对话框
     */
    openFontChangeDialog: function () {
        var url = window.location.href.replace(/index\.html.*$/, "ui/font_change.html");
        if (typeof wps !== "undefined" && wps.ShowDialog) {
            wps.ShowDialog(url, "公文字体批量替换与西文字体设置", 460, 380, false);
        } else {
            window.open(url, "_blank", "width=460,height=380");
        }
    }
};
