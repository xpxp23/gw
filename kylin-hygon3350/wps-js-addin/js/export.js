/**
 * 副本生成与导出模块 (移植自 Ribbon1.Export.cs)
 * 支持在银河麒麟下生成同名 PDF / DOC / PNG / JPG / 仿真扫描版 PDF 副本
 */
var GwExport = {
    /**
     * 确保文档已保存
     */
    ensureDocumentSaved: function (title) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return false;
        var doc = app.ActiveDocument;
        if (doc.Path && doc.Path.length > 0) return true;
        if (!confirm("【" + title + "】\n生成副本前需要先保存当前文档，是否立即保存？")) {
            return false;
        }
        try {
            doc.Save();
            return !!(doc.Path && doc.Path.length > 0);
        } catch (e) {
            alert("保存文档失败，生成副本操作已取消: " + e.message);
            return false;
        }
    },

    getBasePathInfo: function () {
        var app = getWpsApp();
        var doc = app.ActiveDocument;
        var dir = doc.Path.replace(/[\/\\]+$/, "");
        var fullName = doc.Name;
        var baseName = fullName.replace(/\.[^.]+$/, "");
        var sep = dir.indexOf("\\") >= 0 ? "\\" : "/";
        return {
            dir: dir,
            sep: sep,
            fullName: fullName,
            baseName: baseName,
            fullPath: dir + sep + fullName
        };
    },

    /**
     * 生成同名 PDF 副本
     */
    saveAsPdfCopy: function () {
        if (!this.ensureDocumentSaved("生成 PDF 副本")) return;
        var app = getWpsApp();
        var doc = app.ActiveDocument;
        var info = this.getBasePathInfo();
        var outPdf = info.dir + info.sep + info.baseName + ".pdf";
        try {
            app.StatusBar = "正在生成 PDF 格式副本: " + outPdf;
            if (typeof doc.ExportAsFixedFormat === "function") {
                doc.ExportAsFixedFormat(outPdf, WpsEnum.wdExportFormatPDF, true);
            } else {
                doc.SaveAs2(outPdf, WpsEnum.wdFormatPDF);
            }
            app.StatusBar = "已生成 PDF 格式副本: " + outPdf;
        } catch (e) {
            alert("生成 PDF 副本出错: " + e.message);
        }
    },

    /**
     * 生成同名 Word 97-2003 (.doc) 副本
     */
    saveAsDocCopy: function () {
        if (!this.ensureDocumentSaved("生成 DOC 副本")) return;
        var app = getWpsApp();
        var doc = app.ActiveDocument;
        var info = this.getBasePathInfo();
        var outDoc = info.dir + info.sep + info.baseName + ".doc";
        try {
            doc.Save();
            app.StatusBar = "正在生成 Word 97-2003 (.doc) 格式副本...";
            doc.SaveAs2(outDoc, WpsEnum.wdFormatDocument97);
            app.StatusBar = "已生成 DOC 副本: " + outDoc;
            app.Documents.Open(info.fullPath);
        } catch (e) {
            alert("生成 DOC 副本出错: " + e.message);
        }
    },

    /**
     * 生成逐页 PNG / JPG 高清图片副本
     * 在银河麒麟 Linux 下：先导出临时 PDF，再调用本地麒麟桥接服务 (pdftoppm / Pillow) 渲染各页为 1240×1754 高清图片
     */
    saveAsImageCopies: function (imgFormat) {
        if (!this.ensureDocumentSaved("生成 " + imgFormat.toUpperCase() + " 图片副本")) return;
        var app = getWpsApp();
        var doc = app.ActiveDocument;
        var info = this.getBasePathInfo();
        var cfg = GwConfig.load();
        var tempPdf = info.dir + info.sep + "." + info.baseName + "_tmp_export.pdf";

        try {
            app.StatusBar = "正在导出页面渲染源数据...";
            if (typeof doc.ExportAsFixedFormat === "function") {
                doc.ExportAsFixedFormat(tempPdf, WpsEnum.wdExportFormatPDF, false);
            } else {
                doc.SaveAs2(tempPdf, WpsEnum.wdFormatPDF);
            }
        } catch (e) {
            alert("导出页面渲染源失败: " + e.message);
            return;
        }

        GwConfig.callBridge("/api/pdf-to-images", {
            pdfPath: tempPdf,
            outputDir: info.dir,
            baseName: info.baseName,
            format: imgFormat.toLowerCase(),
            dpi: cfg.dpi || 200,
            removeTempPdf: true
        }, function (err, resp) {
            if (err || !resp || !resp.ok) {
                alert("提示：逐页转图片需要本地麒麟桥接服务运行 (xkl-gw-bridge)。\n已为您在同目录生成 PDF 文件：\n" + tempPdf);
            } else {
                app.StatusBar = "已生成 " + resp.count + " 张 " + imgFormat.toUpperCase() + " 页面图片至: " + info.dir;
                alert("导出完成！共生成 " + resp.count + " 张 " + imgFormat.toUpperCase() + " 图片，保存在文档同目录下。");
            }
        });
    },

    /**
     * 生成仿真扫描版 PDF 副本（移植自《生成扫描版pdf.vba》，支持倾斜、噪点、纸张泛黄与灰度化，无需 Windows make-look-scanned.exe）
     */
    saveAsScannedPdfCopy: function () {
        if (!this.ensureDocumentSaved("生成仿真扫描版 PDF")) return;
        var app = getWpsApp();
        var doc = app.ActiveDocument;
        var info = this.getBasePathInfo();
        var tempPdf = info.dir + info.sep + "." + info.baseName + "_tmp_scan.pdf";
        var outScannedPdf = info.dir + info.sep + info.baseName + "_扫描版.pdf";

        try {
            if (typeof doc.ExportAsFixedFormat === "function") {
                doc.ExportAsFixedFormat(tempPdf, WpsEnum.wdExportFormatPDF, false);
            } else {
                doc.SaveAs2(tempPdf, WpsEnum.wdFormatPDF);
            }
        } catch (e) {
            alert("导出基础 PDF 失败: " + e.message);
            return;
        }

        GwConfig.callBridge("/api/make-scanned-pdf", {
            pdfPath: tempPdf,
            outputPath: outScannedPdf,
            skew: 0.6,
            noise: 0.06,
            dpi: 200,
            removeTempPdf: true
        }, function (err, resp) {
            if (err || !resp || !resp.ok) {
                alert("仿真扫描滤镜需要本地麒麟桥接服务支持。\n基础 PDF 已导出至: " + tempPdf);
            } else {
                app.StatusBar = "已生成仿真扫描版 PDF: " + outScannedPdf;
                alert("仿真扫描版 PDF 已生成：\n" + outScannedPdf);
            }
        });
    }
};
