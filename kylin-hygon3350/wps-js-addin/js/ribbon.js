/**
 * 小恐龙公文排版助手 (银河麒麟 V10 / 海光 3350 专版)
 * Ribbon 核心事件分发器与图标映射器 (移植自 Ribbon1.cs)
 */

function OnRibbonLoad(ribbonUI) {
    if (typeof wps !== "undefined") {
        wps.ribbonUI = ribbonUI;
    }
    console.log("小恐龙公文排版助手 (银河麒麟 / 海光3350 专版) Ribbon 加载完成");
}

/**
 * 映射原项目 src/Resources/ 下的全部 PNG 图标资源
 */
function GetImage(control) {
    var id = control && control.Id ? control.Id : "";
    var iconMap = {
        btnPaper: "PageMarginsGallery.png",
        btnInsert2: "AlignJustifyMedium.png",
        btnLeft: "IndentDecrease.png",
        btnHeadDelete: "HeaderFooterEditHeader.png",
        btnHeadNoLine: "HorizontalLineInsert.png",
        glyFormat: "AsianTypography.png",
        btnTitleHei: "TableCellAlignTop.png",
        mnuTitleFontQuick: "FontsReplaceFonts.png",
        btnNormal: "GroupTextBoxText.png",
        mnuTitle1: "_1.png",
        mnuTitle2: "_2.png",
        mnuTitle3: "_3.png",
        mnuTitle4: "_4.png",
        btnUnderDot: "E.png",
        btnHideFont: "H.png",
        glyFont: "FontsReplaceFonts.png",
        btnNum2Txt: "FormattingUnique.png",
        btnHead: "HeaderFooterEditHeader.png",
        btnName: "NewDistributionList.png",
        btnFoot: "HeaderFooterEditFooter.png",
        btnStyle: "ChangeStylesMenu.png",
        mnuInsertGWObject: "OpenOfficialDocumentbubbleImage.png",
        btnHeadSingle: "GWHeadSingle.png",
        btnHeadMulti: "GWHeadMutil.png",
        btnLine1: "GWline1.png",
        btnLineH1: "GWline2.png",
        btnLineH2: "GWline3.png",
        btnBindLine: "CustomPageMargins.png",
        btnFootSingle: "TableCellAlignBottom.png",
        btnFootDouble: "TableCellAlignBottomCenter.png",
        btnDate: "DateAndTimeInsert.png",
        btnInsertPaper: "PageOrientationPortraitLandscape.png",
        glyFoot: "HeaderFooterPageNumberInsert.png",
        btnNumberCenter: "AlignCenter.png",
        btnNumberOutside: "AlignRight.png",
        btnNumberRight: "AlignRight.png",
        btnNumberOnLocation: "HeaderFooterPageNumberInsert.png",
        glySymbol: "StarUnratedFull.png",
        btnInsertSignPic: "StampTool.png",
        btnInsertSignPicSave: "SaveAttachments.png",
        btnInsertSignPicLink: "EditLinks.png",
        btnStampEveryPage: "StampTool.png",
        glyBak: "FileSaveAs.png",
        btnSaveAsPDF: "FileEmailAsPdfEmailAttachment.png",
        btnSaveAsDoc: "FileSaveAsWord97_2003.png",
        btnSaveAsPic: "SaveAsPng.png",
        btnSaveAsJPG: "SaveAsJpg.png",
        btnSaveAsScannedPDF: "PicturePropertiesAlternativeText.png",
        btnProtectDoc: "ProtectDocument.png",
        glyPassword: "PasswordProtectSection.png",
        btnSetPW: "WebControlPassword.png",
        btnClearOpenPW: "SetDatabasePassword.png",
        btnClearProtectPW: "HighlightClear.png",
        mnuSpeakGroup: "SpeakCells.png",
        btnSpeak: "SpeakCells.png",
        btnSpeakOption: "GroupAudioVideoRecordingOptions.png",
        btnVBA: "OfficialTemplate.png",
        mnuQuickScripts: "LogicBoldItalicAction.png",
        mnuSelect: "ContentControlPicture.png",
        btnAbout: "Info.png",
        btnDesktopApp: "OfficialTemplate.png",
        btnHelp: "Help.png"
    };
    var file = iconMap[id] || "OfficialTemplate.png";
    return "images/" + file;
}

/**
 * 1~4 级标题下拉菜单回调
 */
function OnTitleAction(control) {
    var tag = (control && control.Tag) ? control.Tag : "1|1";
    var parts = tag.split("|");
    var level = parseInt(parts[0], 10) || 1;
    var index = parseInt(parts[1], 10) || 1;
    GwFonts.setHeadingLevel(level, index);
}

/**
 * 常用符号插入回调
 */
function OnSymbolAction(control) {
    var sym = (control && control.Tag) ? control.Tag : "";
    if (sym) {
        GwInsert.insertSymbolText(sym);
    }
}

/**
 * 快捷脚本菜单回调
 */
function OnQuickScript(control) {
    var key = (control && control.Tag) ? control.Tag : "";
    if (key) {
        GwVbaPorted.executeByKey(key);
    }
}

/**
 * 主功能按钮点击分发器
 */
function OnAction(control) {
    var id = control && control.Id ? control.Id : "";
    switch (id) {
        // 1. 版面组
        case "btnPaper":
            GwLayout.openPaperSetupDialog();
            break;
        case "btnInsert2":
            GwLayout.indent2Chars();
            break;
        case "btnLeft":
            GwLayout.paragraphFlushLeft();
            break;
        case "btnHeadDelete":
            GwLayout.deletePageHeader();
            break;
        case "btnHeadNoLine":
            GwLayout.removeHeaderLine();
            break;
        case "btnFormatAuto2026":
            GwVbaPorted.runAutoLayout2026();
            break;
        case "btnFormatFullLine":
            GwFormat.fullLineRearrange();
            break;
        case "btnFormatR2N":
            GwFormat.softReturnToParagraph();
            break;
        case "btnFormatDelSpace":
            GwFormat.deleteAllSpaces();
            break;
        case "btnFormatDelSelectSpace":
            GwFormat.deleteHalfSpaces();
            break;
        case "btnFormatDelEmptyLine":
            GwFormat.deleteEmptyLines(false);
            break;
        case "btnFormatDelSelectEmptyLine":
            GwFormat.deleteEmptyLines(true);
            break;
        case "btnFormatDelGraphic":
            GwFormat.deleteGraphics();
            break;
        case "btnFormatRemoveURL":
            GwFormat.removeHyperlinks();
            break;
        case "btnFormatAddSpace":
            GwFormat.addSpaceBetweenCnEn();
            break;
        case "btnFormatE2C":
            GwFormat.punctuationEnToCn();
            break;
        case "btnFormatC2E":
            GwFormat.punctuationCnToEn();
            break;
        case "btnFormatFrame2txt":
            GwFormat.textBoxToText();
            break;
        case "btnFormatParaPage":
            GwFormat.paragraphPageBreakBefore();
            break;

        // 2. 公文标题组
        case "btnTitleHei":
            GwFonts.setMainTitle();
            break;
        case "btnTitleFont_FZXBS":
            GwFonts.setMainTitle("方正小标宋简体");
            break;
        case "btnTitleFont_FZXBSGBK":
            GwFonts.setMainTitle("方正小标宋_GBK");
            break;
        case "btnTitleFont_Hei":
            GwFonts.setMainTitle("黑体");
            break;
        case "btnTitleFont_Song":
            GwFonts.setMainTitle("宋体");
            break;
        case "btnTitleFont_Kai":
            GwFonts.setMainTitle("楷体");
            break;
        case "btnTitleFont_FangSong":
            GwFonts.setMainTitle("仿宋");
            break;

        // 3. 文字格式组
        case "btnNormal":
            GwFonts.setNormalBody();
            break;
        case "btnUnderDot":
            GwFonts.toggleUnderDot();
            break;
        case "btnHideFont":
            GwFonts.toggleHiddenFont();
            break;
        case "btnFontFSgb2312":
            GwFonts.switchFont("仿宋_GB2312", "仿宋");
            break;
        case "btnFontFS":
            GwFonts.switchFont("仿宋", "仿宋_GB2312");
            break;
        case "btnFontKTgb2312":
            GwFonts.switchFont("楷体_GB2312", "楷体");
            break;
        case "btnFontKT":
            GwFonts.switchFont("楷体", "楷体_GB2312");
            break;
        case "btnFontBS2HT":
            GwFonts.switchFont("方正小标宋_GBK", "黑体");
            GwFonts.switchFont("方正小标宋简体", "黑体");
            break;
        case "btnFontHT2BS":
            GwFonts.switchFont("黑体", "方正小标宋_GBK");
            break;
        case "btnFontChangeDialog":
            GwFonts.openFontChangeDialog();
            break;
        case "btnNum2Txt":
            GwLayout.convertNumbersToText();
            break;
        case "btnHead":
            GwFonts.setQuickRoleFont("head");
            break;
        case "btnName":
            GwFonts.setQuickRoleFont("name");
            break;
        case "btnFoot":
            GwFonts.setQuickRoleFont("foot");
            break;
        case "btnStyleCreate":
            GwStyles.createCustomStyles(false);
            break;
        case "btnStyleR":
            GwStyles.createCustomStyles(true);
            break;

        // 4. 插入组
        case "btnHeadSingle":
            GwInsert.insertRedHeadSingle();
            break;
        case "btnHeadMulti":
            GwInsert.insertRedHeadMulti();
            break;
        case "btnLine1":
            GwInsert.insertRedDividerLine();
            break;
        case "btnLineH1":
            GwInsert.insertLetterLine(true);
            break;
        case "btnLineH2":
            GwInsert.insertLetterLine(false);
            break;
        case "btnBindLine":
            GwInsert.insertBindingLine();
            break;
        case "btnFootSingle":
            GwInsert.insertBanjiTable(1);
            break;
        case "btnFootDouble":
            GwInsert.insertBanjiTable(2);
            break;
        case "btnDate":
            GwInsert.insertCnDate();
            break;
        case "btnInsertPaper":
            GwInsert.insertLandscapePage();
            break;
        case "btnNumberCenter":
            GwPageNumber.setPageNumberAlignment(0);
            break;
        case "btnNumberOutside":
            GwPageNumber.setPageNumberAlignment(1);
            break;
        case "btnNumberRight":
            GwPageNumber.setPageNumberAlignment(2);
            break;
        case "btnNumberOnLocation":
            GwPageNumber.insertFieldAtSelection();
            break;
        case "btnSymbol_zkh":
            GwInsert.replaceSquareBracketsToHex();
            break;
        case "btnSymbol_hkh":
            GwInsert.replaceBlackBracketsToHex();
            break;
        case "btnSymbol_index":
            GwInsert.convertNumberToCircled();
            break;
        case "btnInsertSignPicSave":
            GwInsert.insertOfficialStamp(false, false);
            break;
        case "btnInsertSignPicLink":
            GwInsert.insertOfficialStamp(true, false);
            break;
        case "btnStampEveryPage":
            GwInsert.insertOfficialStamp(false, true);
            break;

        // 5. 副本与安全/语音组
        case "btnSaveAsPDF":
            GwExport.saveAsPdfCopy();
            break;
        case "btnSaveAsDoc":
            GwExport.saveAsDocCopy();
            break;
        case "btnSaveAsPic":
            GwExport.saveAsImageCopies("png");
            break;
        case "btnSaveAsJPG":
            GwExport.saveAsImageCopies("jpg");
            break;
        case "btnSaveAsScannedPDF":
            GwExport.saveAsScannedPdfCopy();
            break;
        case "btnProtectDoc":
            GwSecuritySpeech.toggleProtectDoc();
            break;
        case "btnSetPW":
            GwSecuritySpeech.setDocumentPasswords();
            break;
        case "btnClearOpenPW":
            GwSecuritySpeech.clearOpenPassword();
            break;
        case "btnClearProtectPW":
            GwSecuritySpeech.clearWritePassword();
            break;
        case "btnSpeak":
            GwSecuritySpeech.toggleSpeak();
            break;
        case "btnSpeakOption":
            GwSecuritySpeech.openSpeakOptionDialog();
            break;

        // 6. 脚本与对象转换组
        case "btnVBA":
            GwVbaPorted.openVbaCenterDialog();
            break;
        case "btnSelectSharp":
            GwShapes.selectAllShapes();
            break;
        case "btnSelectTables":
            GwShapes.selectAllTables();
            break;
        case "btnSelectCurrentPage":
            GwShapes.selectCurrentPage();
            break;
        case "btnSharp2OnText":
            GwShapes.convertInlineShapes(1);
            break;
        case "btnSharp2Wrapping":
            GwShapes.convertInlineShapes(2);
            break;
        case "btnSharp2TopBonttom":
            GwShapes.convertInlineShapes(3);
            break;
        case "btnSharp2Top":
            GwShapes.convertBehindShapesToFront();
            break;
        case "btnSplitSection":
            GwShapes.splitDocBySection();
            break;
        case "btnOpenDocFolder":
            GwShapes.openDocumentFolderInKylin();
            break;

        // 7. 关于与桌面端联动组
        case "btnAbout":
            var aboutUrl = window.location.href.replace(/index\.html.*$/, "ui/about_settings.html");
            if (typeof wps !== "undefined" && wps.ShowDialog) {
                wps.ShowDialog(aboutUrl, "关于小恐龙公文排版助手 (银河麒麟 / 海光3350 专版)", 540, 480, false);
            } else {
                window.open(aboutUrl, "_blank", "width=540,height=480");
            }
            break;
        case "btnDesktopApp":
            GwConfig.callBridge("/api/launch-desktop-gui", {}, function (err) {
                if (err) {
                    alert("可通过桌面快捷方式【小恐龙公文排版助手(麒麟版)】或终端运行 xkl-gw-gui 启动独立批处理工具。");
                }
            });
            break;
        case "btnHelp":
            var helpUrl = window.location.href.replace(/index\.html.*$/, "ui/vba_center.html");
            if (typeof wps !== "undefined" && wps.ShowDialog) {
                wps.ShowDialog(helpUrl, "公文脚本中心与麒麟版使用说明", 760, 560, false);
            } else {
                window.open(helpUrl, "_blank", "width=760,height=560");
            }
            break;
        default:
            console.log("未知按钮事件:", id);
    }
}
