/**
 * 公文专用样式构建模块 (移植自 Ribbon1.Styles.cs)
 */
var GwStyles = {
    clearCustomStyles: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var styles = app.ActiveDocument.Styles;
            for (var i = styles.Count; i >= 1; i--) {
                var sy = styles.Item(i);
                if (!sy.BuiltIn && sy.NameLocal && sy.NameLocal.indexOf("公文:") === 0) {
                    sy.Delete();
                }
            }
        } catch (e) {
            console.warn("清理自定义样式提示:", e);
        }
    },

    createCustomStyles: function (forceRebuild) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        if (forceRebuild) {
            try { app.Selection.ClearFormatting(); } catch (e) {}
            this.clearCustomStyles();
        }

        var cfg = GwConfig.load();
        var doc = app.ActiveDocument;
        var existing = {};

        try {
            var styles = doc.Styles;
            for (var i = 1; i <= styles.Count; i++) {
                existing[styles.Item(i).NameLocal] = true;
            }
        } catch (e) {}

        function ensureStyle(name, fontSpec, outlineLvl, align, firstIndentChars) {
            if (existing[name]) return;
            try {
                var sy = doc.Styles.Add(name, WpsEnum.wdStyleTypeParagraph);
                sy.Font.NameFarEast = fontSpec.name;
                sy.Font.Name = fontSpec.name;
                sy.Font.Size = parseFloat(fontSpec.size);
                sy.Font.Bold = fontSpec.bold ? 1 : 0;
                sy.Font.Italic = fontSpec.italic ? 1 : 0;
                if (outlineLvl !== null && outlineLvl !== undefined) {
                    sy.ParagraphFormat.OutlineLevel = outlineLvl;
                }
                if (align !== null && align !== undefined) {
                    sy.ParagraphFormat.Alignment = align;
                }
                if (firstIndentChars !== null && firstIndentChars !== undefined) {
                    sy.ParagraphFormat.CharacterUnitFirstLineIndent = firstIndentChars;
                }
            } catch (err) {
                console.warn("创建样式 " + name + " 提示:", err);
            }
        }

        ensureStyle("公文:标题", cfg.fonts.title, WpsEnum.wdOutlineLevelBodyText, WpsEnum.wdAlignParagraphCenter, 0);
        ensureStyle("公文:1级标题", cfg.fonts.title1, WpsEnum.wdOutlineLevel1, WpsEnum.wdAlignParagraphJustify, 2);
        ensureStyle("公文:2级标题", cfg.fonts.title2, WpsEnum.wdOutlineLevel2, WpsEnum.wdAlignParagraphJustify, 2);
        ensureStyle("公文:3级标题", cfg.fonts.title3, WpsEnum.wdOutlineLevel3, WpsEnum.wdAlignParagraphJustify, 2);
        ensureStyle("公文:4级标题", cfg.fonts.title4, WpsEnum.wdOutlineLevel4, WpsEnum.wdAlignParagraphJustify, 2);
        ensureStyle("公文:正文", cfg.fonts.normal, WpsEnum.wdOutlineLevelBodyText, WpsEnum.wdAlignParagraphJustify, 0);
        ensureStyle("公文:正文(缩进)", cfg.fonts.normal, WpsEnum.wdOutlineLevelBodyText, WpsEnum.wdAlignParagraphJustify, 2);

        app.StatusBar = forceRebuild
            ? "已强制重建全套《党政机关公文》专用样式（公文:标题 / 1~4级标题 / 正文 / 正文(缩进)）"
            : "已生成全套《党政机关公文》专用样式";
    }
};
