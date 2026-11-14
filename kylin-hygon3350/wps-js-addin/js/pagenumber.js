/**
 * 公文页码设置模块 (移植自 Ribbon1.PageNumber.cs)
 * 严格遵循 GB/T 9704—2012：四号宋体（— 1 —），单面居中、双面奇右偶左（各空1字）
 */
var GwPageNumber = {
    /**
     * 在给定 Range 或当前 Selection 插入 "— PAGE —" 页码域
     */
    insertFieldAtSelection: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var cfg = GwConfig.load();
        var fSpec = cfg.fonts.pageNumber;
        try {
            app.Selection.Font.NameFarEast = fSpec.name;
            app.Selection.Font.Name = fSpec.name;
            app.Selection.Font.Size = parseFloat(fSpec.size);
            app.Selection.Font.Bold = fSpec.bold ? 1 : 0;
            app.Selection.Font.Italic = fSpec.italic ? 1 : 0;
            app.Selection.TypeText("— ");
            app.Selection.Fields.Add(app.Selection.Range, WpsEnum.wdFieldEmpty, "PAGE \\* Arabic");
            app.Selection.TypeText(" —");
        } catch (e) {
            alert("插入页码域出错: " + e.message);
        }
    },

    /**
     * 设置页码位置:
     * position = 0: 页脚居中 (— 1 —)
     * position = 1: 页脚外侧 (奇数页居右缩进1字符，偶数页居左缩进1字符)
     * position = 2: 页脚居右 (— 1 —)
     */
    setPageNumberAlignment: function (position) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        var cfg = GwConfig.load();
        var fSpec = cfg.fonts.pageNumber;

        try {
            var paneView = app.ActiveWindow.ActivePane.View;
            if (paneView.Type === WpsEnum.wdNormalView || paneView.Type === WpsEnum.wdOutlineView) {
                paneView.Type = WpsEnum.wdPrintView;
            }
            app.ScreenUpdating = false;

            if (position === 0 || position === 2) {
                var targetAlign = (position === 0) ? WpsEnum.wdAlignParagraphCenter : WpsEnum.wdAlignParagraphRight;
                for (var i = 1; i <= doc.Sections.Count; i++) {
                    var s = doc.Sections.Item(i);
                    s.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                    s.Footers.Item(WpsEnum.wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = targetAlign;
                    s.Footers.Item(WpsEnum.wdHeaderFooterEvenPages).Range.ParagraphFormat.Alignment = targetAlign;
                    s.Footers.Item(WpsEnum.wdHeaderFooterFirstPage).Range.ParagraphFormat.Alignment = targetAlign;
                }
                paneView.SeekView = WpsEnum.wdSeekCurrentPageFooter;
                app.Selection.WholeStory();
                app.Selection.Delete();
                this.insertFieldAtSelection();
                app.Selection.ParagraphFormat.Alignment = targetAlign;
                paneView.SeekView = WpsEnum.wdSeekMainDocument;
                app.StatusBar = position === 0 ? "已插入居中公文页码 (— 1 —)" : "已插入居右公文页码 (— 1 —)";
            } else if (position === 1) {
                // 奇偶页不同外侧页码（奇数页居右空1字，偶数页居左空1字）
                doc.PageSetup.OddAndEvenPagesHeaderFooter = -1;
                paneView.SeekView = WpsEnum.wdSeekPrimaryFooter;
                app.Selection.WholeStory();
                app.Selection.Delete();
                this.insertFieldAtSelection();
                app.Selection.ParagraphFormat.CharacterUnitRightIndent = 1;
                app.Selection.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphRight;

                paneView.SeekView = WpsEnum.wdSeekEvenPagesFooter;
                app.Selection.WholeStory();
                app.Selection.Delete();
                this.insertFieldAtSelection();
                app.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 1;
                app.Selection.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphLeft;

                paneView.SeekView = WpsEnum.wdSeekMainDocument;
                app.StatusBar = "已设置奇偶页外侧公文页码（奇右偶左，各空1字）";
            }
        } catch (err) {
            alert("设置公文页码时出错: " + err.message);
        } finally {
            app.ScreenUpdating = true;
        }
    }
};
