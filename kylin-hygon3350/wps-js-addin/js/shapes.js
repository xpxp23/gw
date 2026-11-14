/**
 * 图形、表格选择与环绕转换、分节拆分与麒麟文件管理器联动模块 (移植自 Ribbon1.Shapes.cs)
 */
var GwShapes = {
    /**
     * 选择当前页内容
     */
    selectCurrentPage: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            app.ActiveWindow.ActivePane.View.SeekView = WpsEnum.wdSeekMainDocument;
            app.ActiveDocument.Bookmarks.Item("\\page").Range.Select();
            app.StatusBar = "已选中当前页";
        } catch (e) {
            alert("选择整页时出错: " + e.message);
        }
    },

    /**
     * 选择所有浮动图形
     */
    selectAllShapes: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            app.ActiveDocument.Shapes.SelectAll();
            app.StatusBar = "已选中文档内所有浮动图形";
        } catch (e) {
            alert("选择浮动图形出错: " + e.message);
        }
    },

    /**
     * 选择所有表格
     */
    selectAllTables: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var doc = app.ActiveDocument;
            if (doc.Tables.Count === 0) {
                alert("当前文档中没有表格！");
                return;
            }
            doc.Tables.Item(1).Range.Select();
            app.StatusBar = "当前文档共有 " + doc.Tables.Count + " 个表格，已定位到第 1 个表格";
        } catch (e) {
            alert("选择表格出错: " + e.message);
        }
    },

    /**
     * 批量转换所有嵌入图片(InlineShapes)为浮动图形并指定环绕方式
     * wrapMode: 1=浮于文字上方, 2=四周型, 3=上下型
     */
    convertInlineShapes: function (wrapMode) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        try {
            var inShapes = app.ActiveDocument.InlineShapes;
            var count = 0;
            for (var i = inShapes.Count; i >= 1; i--) {
                var sh = inShapes.Item(i).ConvertToShape();
                if (wrapMode === 1) {
                    sh.WrapFormat.Type = WpsEnum.wdWrapFront;
                } else if (wrapMode === 2) {
                    sh.WrapFormat.Type = WpsEnum.wdWrapSquare;
                } else if (wrapMode === 3) {
                    sh.WrapFormat.Type = WpsEnum.wdWrapTopBottom;
                }
                count++;
            }
            app.StatusBar = "已批量转换 " + count + " 个嵌入图形！";
        } catch (e) {
            alert("转换图片环绕方式出错: " + e.message);
        }
    },

    /**
     * 将所有“衬于文字下方”的图形转为“浮于文字上方”
     */
    convertBehindShapesToFront: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var count = 0;
        try {
            var shapes = app.ActiveDocument.Shapes;
            for (var i = 1; i <= shapes.Count; i++) {
                var sh = shapes.Item(i);
                if (sh.WrapFormat.Type === WpsEnum.wdWrapBehind) {
                    sh.WrapFormat.Type = WpsEnum.wdWrapFront;
                    count++;
                }
            }
            alert("共转换 " + count + " 个衬于文字下方的图形为浮于文字上方！");
        } catch (e) {
            alert("转换衬于文字下方的图形出错: " + e.message);
        }
    },

    /**
     * 按分节符拆分文档
     */
    splitDocBySection: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var oldDoc = app.ActiveDocument;
        if (oldDoc.Sections.Count <= 1) {
            alert("拆分文档前请先插入分节符！（当前文档仅有 1 节）");
            return;
        }
        try {
            var secCount = oldDoc.Sections.Count;
            for (var i = 1; i <= secCount; i++) {
                oldDoc.Sections.Item(i).Range.Copy();
                var newDoc = app.Documents.Add();
                newDoc.Content.Paste();
                GwFormat.wordReplace(false, "^b", "", false);
            }
            app.StatusBar = "已按分节符拆分为 " + secCount + " 个新文档窗口";
        } catch (e) {
            alert("按分节符拆分文档出错: " + e.message);
        }
    },

    /**
     * 在银河麒麟文件管理器 (Peony / xdg-open) 中打开当前文档所在目录
     * 替代原 Windows 版的 explorer.exe /select
     */
    openDocumentFolderInKylin: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var docPath = app.ActiveDocument.FullName || app.ActiveDocument.Path;
        if (!docPath || docPath.indexOf("/") === -1) {
            alert("当前文档尚未保存到本地磁盘，请先保存文档！");
            return;
        }
        GwConfig.callBridge("/api/open-folder", { path: docPath }, function (err, resp) {
            if (err) {
                alert("文档所在路径：\n" + docPath + "\n\n(提示：启动本地麒麟桥接服务后可自动唤起 Peony 文件管理器并高亮选中该文件)");
            } else {
                app.StatusBar = "已在银河麒麟文件管理器 (Peony) 中打开所在文件夹";
            }
        });
    }
};
