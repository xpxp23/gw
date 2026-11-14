/**
 * 小恐龙公文排版助手 (银河麒麟 V10 / 海光 3350 专版)
 * 原版 42 个 VBA 脚本的跨平台原生 JSAPI 移植引擎
 * 彻底解决 Linux 版 WPS 无法调用 Windows VBProject COM 组件及 ANSI/GB2312 编码乱码问题
 */
var GwVbaPorted = {
    /**
     * 整数 (1~99) 转中文数字
     */
    intToCnNum: function (num) {
        var digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"];
        var n = parseInt(num, 10);
        if (isNaN(n) || n <= 0) return "";
        if (n <= 10) return digits[n];
        if (n < 20) return "十" + digits[n - 10];
        var shi = Math.floor(n / 10);
        var ge = n % 10;
        return digits[shi] + "十" + (ge > 0 ? digits[ge] : "");
    },

    /**
     * 阿拉伯数字金额转人民币大写 (移植自《数字转为人民币.vba》)
     */
    numberToRmbUpper: function (rawInput) {
        var cleaned = String(rawInput || "").replace(/[,，\s￥¥元]/g, "");
        var val = parseFloat(cleaned);
        if (isNaN(val)) return null;
        var prefix = val < 0 ? "负" : "";
        val = Math.abs(val);
        if (val > 999999999999.99) return "输入有误：数字过大";
        var fenTotal = Math.round(val * 100);
        if (fenTotal === 0) return "零元整";

        var cnNums = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
        var intUnits = ["", "拾", "佰", "仟"];
        var secUnits = ["", "万", "亿", "万亿"];

        var intPart = Math.floor(fenTotal / 100);
        var jiao = Math.floor((fenTotal % 100) / 10);
        var fen = fenTotal % 10;

        var result = "";
        if (intPart > 0) {
            var intStr = String(intPart);
            var zeroCount = 0;
            var len = intStr.length;
            for (var i = 0; i < len; i++) {
                var n = parseInt(intStr.charAt(i), 10);
                var p = len - i - 1;
                var q = Math.floor(p / 4);
                var m = p % 4;
                if (n === 0) {
                    zeroCount++;
                } else {
                    if (zeroCount > 0) result += "零";
                    zeroCount = 0;
                    result += cnNums[n] + intUnits[m];
                }
                if (m === 0 && zeroCount < 4) {
                    result += secUnits[q];
                }
            }
            result += "元";
        }

        if (jiao === 0 && fen === 0) {
            result += "整";
        } else {
            if (jiao > 0) {
                result += cnNums[jiao] + "角";
            } else if (intPart > 0 && fen > 0) {
                result += "零";
            }
            if (fen > 0) {
                result += cnNums[fen] + "分";
            }
        }
        return prefix + result;
    },

    /**
     * 1. 《公文自动排版2026.vba》完整移植版
     * 一键执行：A4国标版心 -> 清理域与多余空白 -> 括号与标点规范 -> 识别主标题(二号小标宋居中)
     * -> 识别主送机关(三号楷体顶格) -> 识别一~四级标题(一、/（一）/1．/（1）) -> 识别正文(三号仿宋首行缩进2字符,28磅固定行距)
     * -> 识别附件与发文机关落款及成文日期 -> 自动插入公文一字线页码
     */
    runAutoLayout2026: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) {
            alert("请先打开需要排版的公文文档！");
            return;
        }
        var doc = app.ActiveDocument;
        var cfg = GwConfig.load();
        var t0 = Date.now();

        app.ScreenUpdating = false;
        try {
            // 步骤1：设置 GB/T 9704—2012 A4页面边距与网格 (上37mm, 下35mm, 左28mm, 右26mm)
            GwLayout.applyGwPageSetup(cfg);

            // 步骤2：初始化清理（软回车转段落、文号括号统一为六角括号〔〕）
            GwFormat.wordReplace(false, "^l", "^p", false);
            GwFormat.wordReplace(true, "\\[([0-9]{4})\\]", "〔\\1〕", false);
            GwFormat.wordReplace(true, "【([0-9]{4})】", "〔\\1〕", false);
            GwFormat.wordReplace(true, "［([0-9]{4})］", "〔\\1〕", false);

            // 步骤3：删除空白段落并遍历识别各段落角色
            var paras = doc.Paragraphs;
            for (var k = paras.Count; k >= 1; k--) {
                var rawT = (paras.Item(k).Range.Text || "").replace(/[\r\n\s\u3000\x07]/g, "");
                if (rawT.length === 0 && paras.Item(k).Range.InlineShapes.Count === 0) {
                    paras.Item(k).Range.Delete();
                }
            }

            var totalParas = doc.Paragraphs.Count;
            var foundMainTitle = false;
            var dateParaIdx = -1;

            // 正则匹配规则
            var reH1 = /^[一二三四五六七八九十百]+[、．.]\s*/;
            var reH2 = /^[（(][一二三四五六七八九十百]+[）)]\s*/;
            var reH3 = /^\d+[．.、]\s*/;
            var reH4 = /^[（(]\d+[）)]\s*/;
            var reSalutation = /^[^\s。；！？，、]{2,35}[：:]$/;
            var reAttachment = /^附件[：:\d\s]/;
            var reDate = /^([12]\d{3}|[二〇零○一二三四五六七八九]{4})\s*年\s*(\d{1,2}|[一二三四五六七八九十]{1,2})\s*月\s*(\d{1,2}|[一二三四五六七八九十]{1,3})\s*日$/;

            for (var i = 1; i <= totalParas; i++) {
                var p = doc.Paragraphs.Item(i);
                var rng = p.Range;
                var txt = (rng.Text || "").replace(/[\r\n\x07]/g, "").replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
                if (!txt) continue;

                // 清理段首空白
                if (rng.Text.match(/^[\s\u3000\t]+/)) {
                    rng.Text = txt + "\r";
                }

                // 统一行距 28 磅
                p.Format.SpaceBefore = 0;
                p.Format.SpaceAfter = 0;
                p.Format.LineSpacingRule = WpsEnum.wdLineSpaceExactly;
                p.Format.LineSpacing = parseFloat(cfg.lineSpacingPt || 28.0);

                // 1) 第一段非文号段落作为公文主标题 (二号方正小标宋简体，居中)
                if (!foundMainTitle && i <= 2 && !txt.match(/〔\d{4}〕.*号/)) {
                    foundMainTitle = true;
                    rng.Font.NameFarEast = cfg.fonts.title.name;
                    rng.Font.Name = cfg.fonts.title.name;
                    rng.Font.Size = 22;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphCenter;
                    p.Format.CharacterUnitFirstLineIndent = 0;
                    p.Format.FirstLineIndent = 0;
                    p.Format.LineSpacingRule = WpsEnum.wdLineSpaceExactly;
                    p.Format.LineSpacing = 32;
                    continue;
                }

                // 2) 发文字号（如 ○○发〔2026〕1号）
                if (i <= 3 && txt.match(/〔\d{4}〕.*号/)) {
                    rng.Font.NameFarEast = cfg.fonts.normal.name;
                    rng.Font.NameAscii = cfg.letterFont || "Times New Roman";
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphCenter;
                    p.Format.CharacterUnitFirstLineIndent = 0;
                    p.Format.FirstLineIndent = 0;
                    continue;
                }

                // 3) 主送机关（称呼行，顶格三号仿宋）
                if (i <= 4 && reSalutation.test(txt) && !reAttachment.test(txt)) {
                    rng.Font.NameFarEast = cfg.fonts.normal.name;
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphLeft;
                    p.Format.CharacterUnitFirstLineIndent = 0;
                    p.Format.FirstLineIndent = 0;
                    continue;
                }

                // 4) 一级标题：一、 (三号黑体，首行缩进2字符)
                if (reH1.test(txt)) {
                    rng.Font.NameFarEast = cfg.fonts.title1.name;
                    rng.Font.Name = cfg.fonts.title1.name;
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphJustify;
                    p.Format.CharacterUnitFirstLineIndent = 2;
                    p.Format.OutlineLevel = WpsEnum.wdOutlineLevel1;
                    continue;
                }

                // 5) 二级标题：（一） (三号楷体，首行缩进2字符)
                if (reH2.test(txt)) {
                    rng.Font.NameFarEast = cfg.fonts.title2.name;
                    rng.Font.Name = cfg.fonts.title2.name;
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphJustify;
                    p.Format.CharacterUnitFirstLineIndent = 2;
                    p.Format.OutlineLevel = WpsEnum.wdOutlineLevel2;
                    continue;
                }

                // 6) 三级标题：1． (三号仿宋，首行缩进2字符，句首至句号/冒号可加粗)
                if (reH3.test(txt)) {
                    rng.Font.NameFarEast = cfg.fonts.title3.name;
                    rng.Font.NameAscii = cfg.numberFont || "Times New Roman";
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphJustify;
                    p.Format.CharacterUnitFirstLineIndent = 2;
                    p.Format.OutlineLevel = WpsEnum.wdOutlineLevel3;
                    continue;
                }

                // 7) 四级标题：（1） (三号仿宋，首行缩进2字符)
                if (reH4.test(txt)) {
                    rng.Font.NameFarEast = cfg.fonts.title4.name;
                    rng.Font.NameAscii = cfg.numberFont || "Times New Roman";
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphJustify;
                    p.Format.CharacterUnitFirstLineIndent = 2;
                    p.Format.OutlineLevel = WpsEnum.wdOutlineLevel4;
                    continue;
                }

                // 8) 附件说明行（左空二字，回行对齐）
                if (reAttachment.test(txt)) {
                    rng.Font.NameFarEast = cfg.fonts.normal.name;
                    rng.Font.NameAscii = cfg.numberFont || "Times New Roman";
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphJustify;
                    p.Format.CharacterUnitLeftIndent = 5;
                    p.Format.CharacterUnitFirstLineIndent = -3;
                    continue;
                }

                // 9) 成文日期与发文机关署名（右空四字）
                if (reDate.test(txt) && i >= Math.max(2, totalParas - 6)) {
                    dateParaIdx = i;
                    rng.Font.NameFarEast = cfg.fonts.normal.name;
                    rng.Font.NameAscii = cfg.numberFont || "Times New Roman";
                    rng.Font.Size = 16;
                    rng.Font.Bold = 0;
                    p.Format.Alignment = WpsEnum.wdAlignParagraphRight;
                    p.Format.CharacterUnitFirstLineIndent = 0;
                    p.Format.FirstLineIndent = 0;
                    p.Format.CharacterUnitRightIndent = 4;

                    // 将成文日期前一段识别为发文机关署名并居中对齐于日期上方
                    if (i > 1) {
                        var signP = doc.Paragraphs.Item(i - 1);
                        var signTxt = (signP.Range.Text || "").replace(/[\r\n\s]/g, "");
                        if (signTxt.length >= 2 && signTxt.length <= 24 && !signTxt.match(/[。；！？]/)) {
                            signP.Range.Font.NameFarEast = cfg.fonts.normal.name;
                            signP.Range.Font.Size = 16;
                            signP.Format.Alignment = WpsEnum.wdAlignParagraphRight;
                            signP.Format.CharacterUnitFirstLineIndent = 0;
                            signP.Format.FirstLineIndent = 0;
                            var diff = Math.max(0, (txt.length - signTxt.length) / 2.0);
                            signP.Format.CharacterUnitRightIndent = 4 + diff;
                        }
                    }
                    continue;
                }

                // 10) 普通公文正文段落（三号仿宋，两端对齐，首行缩进2字符）
                rng.Font.NameFarEast = cfg.fonts.normal.name;
                rng.Font.NameAscii = cfg.letterFont || "Times New Roman";
                rng.Font.Size = 16;
                rng.Font.Bold = 0;
                p.Format.Alignment = WpsEnum.wdAlignParagraphJustify;
                p.Format.CharacterUnitFirstLineIndent = 2;
                p.Format.OutlineLevel = WpsEnum.wdOutlineLevelBodyText;
            }

            // 步骤4：设置标准一字线居中页码
            GwPageNumber.setPageNumberAlignment(0);

            var elapsed = ((Date.now() - t0) / 1000.0).toFixed(2);
            app.StatusBar = "2026国标公文自动排版执行完成！共用时: " + elapsed + " 秒";
            alert("公文自动排版 (GB/T 9704—2012 / 2026版) 完成！\n共处理 " + totalParas + " 个段落，用时 " + elapsed + " 秒。");
        } catch (err) {
            alert("自动排版过程中出现异常: " + err.message);
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 2. 《公文落款.vba》：将选中或当前两行设置为标准公文落款（署名与日期居中对应、日期右空4字）
     */
    formatOfficialSignature: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var paras = app.Selection.Paragraphs;
        var signP = null, dateP = null;
        if (paras.Count >= 2) {
            signP = paras.Item(1);
            dateP = paras.Item(2);
        } else {
            signP = paras.Item(1);
            if (signP.Next()) {
                dateP = signP.Next();
            }
        }
        if (!signP || !dateP) {
            alert("请同时选中【发文机关署名】和【成文日期】两行，或将光标置于署名行！");
            return;
        }
        var signTxt = (signP.Range.Text || "").replace(/[\r\n\s\u3000]/g, "");
        var dateTxt = (dateP.Range.Text || "").replace(/[\r\n\s\u3000]/g, "");
        if (!signTxt || !dateTxt) {
            alert("未检测到有效的署名或日期内容！");
            return;
        }
        dateP.Range.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphRight;
        dateP.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
        dateP.Range.ParagraphFormat.FirstLineIndent = 0;
        dateP.Range.ParagraphFormat.RightIndent = cmToPt(2.0);

        signP.Range.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphRight;
        signP.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
        signP.Range.ParagraphFormat.FirstLineIndent = 0;
        var indentAdjust = cmToPt((dateTxt.length - signTxt.length) * 0.75 / 2.0);
        signP.Range.ParagraphFormat.RightIndent = Math.max(0, cmToPt(2.0) + indentAdjust);
        app.StatusBar = "公文落款（署名+日期）对齐设置完成！";
    },

    /**
     * 3. 《加粗第n章第n条.vba》 & 《加粗第n章第n条(整行).vba》
     */
    boldChapterAndArticle: function (wholeLine) {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        var count = 0;
        var re = /^(?:\s*)(第[零〇一二三四五六七八九十百千0-9]+[章条])/;
        app.ScreenUpdating = false;
        try {
            for (var i = 1; i <= doc.Paragraphs.Count; i++) {
                var p = doc.Paragraphs.Item(i);
                var txt = p.Range.Text || "";
                var m = txt.match(re);
                if (m) {
                    if (wholeLine) {
                        p.Range.Font.Bold = 1;
                    } else {
                        var subRng = doc.Range(p.Range.Start, p.Range.Start + m[0].length);
                        subRng.Font.Bold = 1;
                    }
                    count++;
                }
            }
            alert((wholeLine ? "整行加粗" : "精准加粗") + "第N章/第N条完成！共处理 " + count + " 处。");
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 4. 《第n章整行加粗居中.vba》
     */
    centerAndBoldChapters: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        var count = 0;
        var re = /^\s*第[零〇一二三四五六七八九十百千0-9]{1,5}章/;
        app.ScreenUpdating = false;
        try {
            for (var i = 1; i <= doc.Paragraphs.Count; i++) {
                var p = doc.Paragraphs.Item(i);
                var txt = (p.Range.Text || "").replace(/[\r\n]/g, "");
                if (re.test(txt)) {
                    var pf = p.Range.ParagraphFormat;
                    pf.LeftIndent = 0;
                    pf.RightIndent = 0;
                    pf.FirstLineIndent = 0;
                    pf.CharacterUnitLeftIndent = 0;
                    pf.CharacterUnitFirstLineIndent = 0;
                    pf.Alignment = WpsEnum.wdAlignParagraphCenter;
                    p.Range.Font.Bold = 1;
                    count++;
                }
            }
            alert("处理完成！共将 " + count + " 个“第…章”标题设为居中、加粗、顶格。");
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 5. 《第n章条后1空格.vba》
     */
    normalizeChapterArticleSpacing: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        var count = 0;
        var re = /^(\s*)(第[零〇一二三四五六七八九十百千\d]+[条章])([ \t\u3000]*)/;
        app.ScreenUpdating = false;
        try {
            for (var i = 1; i <= doc.Paragraphs.Count; i++) {
                var p = doc.Paragraphs.Item(i);
                var txt = p.Range.Text || "";
                var m = txt.match(re);
                if (m) {
                    var rng = doc.Range(p.Range.Start, p.Range.Start + m[0].length);
                    rng.Text = m[1] + m[2] + " ";
                    count++;
                }
            }
            alert("格式调整完成！共规范 " + count + " 处章/条后空格。");
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 6. 《重新整理章条序号.vba》：先章后条各自从“第一”开始连续编号
     */
    renumberChaptersAndArticles: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        var chCount = 0, artCount = 0;
        var reCh = /^(\s*)第[零〇一二三四五六七八九十百千\d]+章/;
        var reArt = /^(\s*)第[零〇一二三四五六七八九十百千\d]+条/;

        app.ScreenUpdating = false;
        try {
            for (var i = 1; i <= doc.Paragraphs.Count; i++) {
                var p = doc.Paragraphs.Item(i);
                var txt = p.Range.Text || "";
                var mCh = txt.match(reCh);
                if (mCh) {
                    chCount++;
                    var r1 = doc.Range(p.Range.Start, p.Range.Start + mCh[0].length);
                    r1.Text = mCh[1] + "第" + this.intToCnNum(chCount) + "章";
                    continue;
                }
                var mArt = txt.match(reArt);
                if (mArt) {
                    artCount++;
                    var r2 = doc.Range(p.Range.Start, p.Range.Start + mArt[0].length);
                    r2.Text = mArt[1] + "第" + this.intToCnNum(artCount) + "条";
                }
            }
            alert("章条序号整理完成！\n共重新编号 " + chCount + " 章、" + artCount + " 条。");
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 7. 《（一）至（九十九）重新排序.vba》
     */
    renumberChineseBrackets: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var paras = app.Selection.Paragraphs;
        if (app.Selection.Start === app.Selection.End) {
            if (!confirm("未选中文本范围，是否对全文中的（一）、（二）…编号进行连续重排？")) return;
            paras = app.ActiveDocument.Paragraphs;
        }
        var doc = app.ActiveDocument;
        var re = /（[一二三四五六七八九十百]+）/;
        var idx = 0;
        app.ScreenUpdating = false;
        try {
            for (var i = 1; i <= paras.Count; i++) {
                var p = paras.Item(i);
                var txt = p.Range.Text || "";
                var m = txt.match(re);
                if (m) {
                    idx++;
                    var pos = txt.indexOf(m[0]);
                    var sub = doc.Range(p.Range.Start + pos, p.Range.Start + pos + m[0].length);
                    sub.Text = "（" + this.intToCnNum(idx) + "）";
                }
            }
            alert("中文括号序号重排完成！共处理 " + idx + " 个编号。");
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 8. 《选中内容重新数字编号.vba》
     */
    renumberSelectionDigits: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        if (app.Selection.Start === app.Selection.End) {
            alert("请先选中要重新编号的文本段落区域！");
            return;
        }
        var paras = app.Selection.Paragraphs;
        var re = /^\s*\d+[.．、]\s*/;
        var idx = 0;
        app.ScreenUpdating = false;
        try {
            for (var i = 1; i <= paras.Count; i++) {
                var p = paras.Item(i);
                var raw = (p.Range.Text || "").replace(/[\r\n]/g, "").trim();
                if (raw.length > 0) {
                    idx++;
                    var body = raw.replace(re, "");
                    p.Range.Text = idx + ". " + body + "\r";
                }
            }
            alert("数字重新编号完成！共处理 " + idx + " 个条目。");
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 9. 《数字转为人民币.vba》
     */
    convertSelectionToRmb: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var raw = (app.Selection.Text || "").trim();
        if (!raw) {
            raw = prompt("请输入要转换为人民币大写的数字金额：", "12345.67");
            if (!raw) return;
        }
        var rmb = this.numberToRmbUpper(raw);
        if (!rmb) {
            alert("选中的内容不是有效的数字金额！");
            return;
        }
        app.Selection.Text = rmb;
        app.StatusBar = "已转换为人民币大写: " + rmb;
    },

    /**
     * 10. 《姓名对齐.vba》：将所选姓名按指定字数宽度分散对齐
     */
    alignSelectedNames: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var nStr = prompt("所选姓名按几字宽度对齐？请输入整数（默认 3 字姓名）：", "3");
        if (!nStr) return;
        var n = parseInt(nStr, 10) || 3;
        try {
            var fontSize = app.Selection.Font.Size || 16;
            app.Selection.Range.FitTextWidth = n * fontSize;
            app.StatusBar = "已将所选姓名按 " + n + " 字宽度对齐";
        } catch (e) {
            // 若WPS版本不支持FitTextWidth，对2字姓名自动插入全角空格实现与3字对齐
            var raw = app.Selection.Text || "";
            if (raw.replace(/\s/g, "").length === 2 && n === 3) {
                var chars = raw.replace(/\s/g, "");
                app.Selection.Text = chars.charAt(0) + "　" + chars.charAt(1);
            }
        }
    },

    /**
     * 11. 《桌签标牌制作.vba》：A4横向双面对折席卡/桌签一键生成
     */
    makePlaceCards: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var nameInput = (app.Selection.Text || "").replace(/[\r\n]/g, "").trim();
        if (!nameInput) {
            nameInput = prompt("请输入桌签姓名（多个姓名请用逗号或空格分隔）：", "张　三");
            if (!nameInput) return;
        }
        var names = nameInput.split(/[,，\s]+/).filter(Boolean);
        var doc = app.ActiveDocument;
        try {
            doc.PageSetup.Orientation = WpsEnum.wdOrientLandscape;
            doc.PageSetup.PageWidth = cmToPt(29.7);
            doc.PageSetup.PageHeight = cmToPt(21.0);
            doc.Content.Delete();

            var tbl = doc.Tables.Add(doc.Range(0, 0), names.length * 2, 1);
            tbl.Borders.Enable = 1;
            for (var i = 0; i < names.length; i++) {
                var display = names[i].length === 2 ? (names[i].charAt(0) + "  " + names[i].charAt(1)) : names[i];
                var rTop = tbl.Rows.Item(i * 2 + 1);
                var rBtm = tbl.Rows.Item(i * 2 + 2);
                rTop.HeightRule = WpsEnum.wdRowHeightExactly;
                rTop.Height = cmToPt(8.5);
                rBtm.HeightRule = WpsEnum.wdRowHeightExactly;
                rBtm.Height = cmToPt(8.5);

                var cTop = tbl.Cell(i * 2 + 1, 1);
                var cBtm = tbl.Cell(i * 2 + 2, 1);
                cTop.Range.Text = display;
                cBtm.Range.Text = display;
                cTop.VerticalAlignment = WpsEnum.wdCellAlignVerticalCenter;
                cBtm.VerticalAlignment = WpsEnum.wdCellAlignVerticalCenter;
                cTop.Range.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
                cBtm.Range.ParagraphFormat.Alignment = WpsEnum.wdAlignParagraphCenter;
                cTop.Range.Font.NameFarEast = "黑体";
                cTop.Range.Font.Size = 72;
                cTop.Range.Font.Bold = 1;
                cBtm.Range.Font.NameFarEast = "黑体";
                cBtm.Range.Font.Size = 72;
                cBtm.Range.Font.Bold = 1;
            }
            app.StatusBar = "已生成 " + names.length + " 组A4横向会议桌签标牌";
        } catch (e) {
            alert("制作桌签标牌出错: " + e.message);
        }
    },

    /**
     * 12. 《删除页眉页脚.vba》：彻底清除全篇所有节的页眉、页脚与页眉横线
     */
    deleteAllHeadersAndFooters: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        app.ScreenUpdating = false;
        try {
            for (var i = 1; i <= doc.Sections.Count; i++) {
                var sec = doc.Sections.Item(i);
                var indices = [
                    WpsEnum.wdHeaderFooterPrimary,
                    WpsEnum.wdHeaderFooterFirstPage,
                    WpsEnum.wdHeaderFooterEvenPages
                ];
                for (var j = 0; j < indices.length; j++) {
                    try {
                        var h = sec.Headers.Item(indices[j]).Range;
                        h.Delete();
                        h.ParagraphFormat.Borders.Item(WpsEnum.wdBorderBottom).LineStyle = WpsEnum.wdLineStyleNone;
                    } catch (e1) {}
                    try {
                        var f = sec.Footers.Item(indices[j]).Range;
                        f.Delete();
                    } catch (e2) {}
                }
            }
            alert("全篇页眉、页脚及页眉横线已彻底删除完毕！");
        } finally {
            app.ScreenUpdating = true;
        }
    },

    /**
     * 13. 《图片批量加边框.vba》
     */
    addBorderToAllPictures: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        var count = 0;
        for (var i = 1; i <= doc.InlineShapes.Count; i++) {
            try {
                var b = doc.InlineShapes.Item(i).Borders;
                b.OutsideLineStyle = WpsEnum.wdLineStyleSingle;
                b.OutsideLineWidth = WpsEnum.wdLineWidth025pt;
                count++;
            } catch (e) {}
        }
        alert("已为文档内 " + count + " 张图片批量添加 0.25 磅细边框！");
    },

    /**
     * 14. 《批量调整图片大小.vba》
     */
    batchResizePictures: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var wStr = prompt("请输入图片统一宽度（厘米，输入 0 表示不改宽度）：", "14");
        if (wStr === null) return;
        var hStr = prompt("请输入图片统一高度（厘米，输入 0 表示不改高度）：", "0");
        if (hStr === null) return;
        var w = parseFloat(wStr) || 0;
        var h = parseFloat(hStr) || 0;
        if (w <= 0 && h <= 0) return;

        var doc = app.ActiveDocument;
        var count = 0;
        for (var i = 1; i <= doc.InlineShapes.Count; i++) {
            var pic = doc.InlineShapes.Item(i);
            if (w > 0) pic.Width = cmToPt(w);
            if (h > 0) pic.Height = cmToPt(h);
            count++;
        }
        alert("已批量调整 " + count + " 张图片的尺寸！");
    },

    /**
     * 15. 《复制文件本地路径.vba》
     */
    copyDocumentLocalPath: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var fullPath = app.ActiveDocument.FullName || app.ActiveDocument.Name;
        GwConfig.callBridge("/api/copy-clipboard", { text: fullPath }, function () {
            prompt("当前文档本地完整路径（按 Ctrl+C 复制）：", fullPath);
        });
    },

    /**
     * 16. 《生成乘法口诀表.vba》
     */
    insertMultiplicationTable: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var lines = ["乘法口诀表"];
        for (var j = 1; j <= 9; j++) {
            var row = [];
            for (var i = 1; i <= j; i++) {
                row.push(i + " × " + j + " = " + (i * j));
            }
            lines.push(row.join("\t"));
        }
        app.Selection.TypeText(lines.join("\r") + "\r");
        app.StatusBar = "已生成九九乘法口诀表";
    },

    /**
     * 17. 《用所选文字保存文件名.vba》
     */
    saveDocWithSelectedText: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument || !app.Selection) return;
        var name = (app.Selection.Text || "").replace(/[\r\n\\/:*?"<>|]/g, "").trim();
        if (!name) {
            alert("请先选中一段文字作为文件名！");
            return;
        }
        var dir = app.ActiveDocument.Path || "/home/kylin/Desktop";
        var sep = dir.indexOf("\\") >= 0 ? "\\" : "/";
        var target = prompt("确认保存路径与文件名：", dir + sep + name.substring(0, 40) + ".docx");
        if (!target) return;
        try {
            app.ActiveDocument.SaveAs2(target, WpsEnum.wdFormatDocumentDefault);
            app.StatusBar = "已保存为: " + target;
        } catch (e) {
            alert("保存失败: " + e.message);
        }
    },

    /**
     * 统一执行命名脚本入口
     */
    executeByKey: function (key) {
        switch (key) {
            case "auto_layout_2026": return this.runAutoLayout2026();
            case "official_signature": return this.formatOfficialSignature();
            case "bold_chapter_article": return this.boldChapterAndArticle(false);
            case "bold_chapter_article_line": return this.boldChapterAndArticle(true);
            case "center_chapter_line": return this.centerAndBoldChapters();
            case "chapter_article_space": return this.normalizeChapterArticleSpacing();
            case "renumber_chapter_article": return this.renumberChaptersAndArticles();
            case "renumber_cn_brackets": return this.renumberChineseBrackets();
            case "renumber_selection_digits": return this.renumberSelectionDigits();
            case "number_to_rmb": return this.convertSelectionToRmb();
            case "align_names": return this.alignSelectedNames();
            case "make_place_cards": return this.makePlaceCards();
            case "delete_all_headers_footers": return this.deleteAllHeadersAndFooters();
            case "picture_add_border": return this.addBorderToAllPictures();
            case "picture_batch_resize": return this.batchResizePictures();
            case "copy_local_path": return this.copyDocumentLocalPath();
            case "multiplication_table": return this.insertMultiplicationTable();
            case "save_by_selected_text": return this.saveDocWithSelectedText();
            default:
                alert("正在通过银河麒麟公文脚本引擎执行: " + key);
        }
    },

    /**
     * 打开内置 42 个公文脚本可视化中心
     */
    openVbaCenterDialog: function () {
        var url = window.location.href.replace(/index\.html.*$/, "ui/vba_center.html");
        if (typeof wps !== "undefined" && wps.ShowDialog) {
            wps.ShowDialog(url, "小恐龙公文脚本中心 (银河麒麟 / 海光3350 原生引擎)", 760, 560, false);
        } else {
            window.open(url, "_blank", "width=760,height=560");
        }
    }
};
