/**
 * 文字排版与批量替换清理模块 (移植自 Ribbon1.Format.cs)
 */
var GwFormat = {
    /**
     * 通用查找替换辅助函数
     */
    wordReplace: function (wildcard, findString, replaceString, selectionOnly) {
        var app = getWpsApp();
        if (!app || !app.Selection) return false;
        try {
            var findObj = selectionOnly ? app.Selection.Find : app.ActiveDocument.Content.Find;
            findObj.ClearFormatting();
            findObj.Replacement.ClearFormatting();
            var wrapMode = selectionOnly ? WpsEnum.wdFindStop : WpsEnum.wdFindContinue;
            return findObj.Execute(
                findString,
                false,
                false,
                wildcard,
                false,
                false,
                true,
                wrapMode,
                false,
                replaceString,
                WpsEnum.wdReplaceAll
            );
        } catch (e) {
            console.error("wordReplace error:", e);
            return false;
        }
    },

    /**
     * 检查是否有选区；若未选定则询问是否全文操作
     * 返回: { proceed: boolean, selectionOnly: boolean }
     */
    checkScope: function (title) {
        var app = getWpsApp();
        if (!app || !app.Selection) return { proceed: false, selectionOnly: false };
        if (app.Selection.Characters.Count > 1 && app.Selection.Start !== app.Selection.End) {
            return { proceed: true, selectionOnly: true };
        }
        var ok = confirm("【" + title + "】\n您当前未选定文字范围，是否对全文执行此操作？");
        return { proceed: ok, selectionOnly: false };
    },

    /**
     * 断行重排：合并未以句末标点结尾的异常换行
     */
    fullLineRearrange: function () {
        var scope = this.checkScope("断行重排");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            this.wordReplace(true, "([!。：！？；])^13", "\\1", scope.selectionOnly);
            app.StatusBar = "已完成断行重排";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 软回车转硬回车 (^l -> ^p)
     */
    softReturnToParagraph: function () {
        var scope = this.checkScope("换行符转回车");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            this.wordReplace(false, "^l", "^p", scope.selectionOnly);
            app.StatusBar = "已将所有手动换行符转换为段落回车";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 删除所有空白（含半角空格、全角空格、制表符）
     */
    deleteAllSpaces: function () {
        var scope = this.checkScope("删除所有空白");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            this.wordReplace(false, "^w", "", scope.selectionOnly);
            this.wordReplace(false, "　", "", scope.selectionOnly);
            this.wordReplace(false, " ", "", scope.selectionOnly);
            app.StatusBar = "已删除范围内所有空白字符";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 删除半角空格
     */
    deleteHalfSpaces: function () {
        var scope = this.checkScope("删除半角空格");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            this.wordReplace(false, " ", "", scope.selectionOnly);
            app.StatusBar = "已删除半角空格";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 删除空行（全文或选区）
     */
    deleteEmptyLines: function (forceSelectionOnly) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var paras = app.Selection.Paragraphs;
        if (!forceSelectionOnly && app.Selection.Characters.Count <= 1) {
            if (!confirm("【删除空行】\n您未选定范围，是否删除全文中的所有空白段落？")) {
                return;
            }
            paras = app.ActiveDocument.Paragraphs;
        }
        var deletedCount = 0;
        app.ScreenUpdating = false;
        try {
            for (var i = paras.Count; i >= 1; i--) {
                var p = paras.Item(i);
                var txt = (p.Range.Text || "").replace(/[\r\n\s\u3000\x07]/g, "");
                if (txt.length === 0 && p.Range.InlineShapes.Count === 0) {
                    p.Range.Delete();
                    deletedCount++;
                }
            }
            app.StatusBar = "已清理 " + deletedCount + " 个空段落！";
        } catch (e) {
            alert("删除空行时出错: " + e.message);
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 删除文档中所有嵌入图片与图表
     */
    deleteGraphics: function () {
        var scope = this.checkScope("删除图片图表");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            this.wordReplace(false, "^g", "", scope.selectionOnly);
            app.StatusBar = "已删除嵌入图片与图表";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 移除超链接
     */
    removeHyperlinks: function () {
        var scope = this.checkScope("移除超链接");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            var target = scope.selectionOnly ? app.Selection : app.ActiveDocument;
            var hls = target.Hyperlinks;
            for (var i = hls.Count; i >= 1; i--) {
                hls.Item(i).Delete();
            }
            app.StatusBar = "已移除超链接";
        } catch (e) {
            alert("移除超链接出错: " + e.message);
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 中英文之间添加半角空格
     */
    addSpaceBetweenCnEn: function () {
        var scope = this.checkScope("中英文之间加空格");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            this.wordReplace(true, "([A-Za-z])([一-龥])", "\\1 \\2", scope.selectionOnly);
            this.wordReplace(true, "([一-龥])([A-Za-z])", "\\1 \\2", scope.selectionOnly);
            app.StatusBar = "已在中文与英文字母之间添加空格";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 标点符号：英文半角转中文全角
     */
    punctuationEnToCn: function () {
        var scope = this.checkScope("标点符号(英转中)");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            this.wordReplace(true, "([!0-9])\\.([!0-9])", "\\1。\\2", scope.selectionOnly);
            var eArr = [",", ";", ":", "?", "!", "-", "~", "(", ")", "[", "]"];
            var cArr = ["，", "；", "：", "？", "！", "—", "～", "（", "）", "［", "］"];
            for (var i = 0; i < eArr.length; i++) {
                this.wordReplace(false, eArr[i], cArr[i], scope.selectionOnly);
            }
            app.StatusBar = "已将英文标点转换为中文全角标点";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 标点符号：中文全角转英文半角
     */
    punctuationCnToEn: function () {
        var scope = this.checkScope("标点符号(中转英)");
        if (!scope.proceed) return;
        var app = getWpsApp();
        app.ScreenUpdating = false;
        try {
            var cArr = ["。", "，", "；", "：", "？", "！", "—", "～", "（", "）", "［", "］", "“", "”", "‘", "’"];
            var eArr = [".", ",", ";", ":", "?", "!", "-", "~", "(", ")", "[", "]", "\"", "\"", "'", "'"];
            for (var i = 0; i < cArr.length; i++) {
                this.wordReplace(false, cArr[i], eArr[i], scope.selectionOnly);
            }
            app.StatusBar = "已将中文标点转换为英文半角标点";
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 文本框转普通文字
     */
    textBoxToText: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            if (app.Selection.ShapeRange && app.Selection.ShapeRange.Count > 0) {
                var sh = app.Selection.ShapeRange.Item(1);
                var txt = sh.TextFrame.TextRange.Text;
                sh.Delete();
                app.Selection.TypeText(txt);
                app.StatusBar = "已将选中文本框转换为正文文字";
            } else {
                alert("请先选中一个文本框！");
            }
        } catch (e) {
            alert("文本框转文字失败: " + e.message);
        }
    },

    /**
     * 所选段落段前分页
     */
    paragraphPageBreakBefore: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        try {
            app.Selection.ParagraphFormat.PageBreakBefore = -1;
            app.StatusBar = "已设置所选段落段前分页";
        } catch (e) {
            alert("设置段前分页失败: " + e.message);
        }
    }
};
