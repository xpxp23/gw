/**
 * WPS Office Linux (Writer / 文字) 枚举常量兼容层
 * 适用于银河麒麟 V10 + 海光 3350 (x86_64) 下的 WPS Office 2019/2023/2026 专业版与信创版
 */
var WpsEnum = {
    // 段落对齐 WdParagraphAlignment
    wdAlignParagraphLeft: 0,
    wdAlignParagraphCenter: 1,
    wdAlignParagraphRight: 2,
    wdAlignParagraphJustify: 3,
    wdAlignParagraphDistribute: 4,

    // 页面方向 WdOrientation
    wdOrientPortrait: 0,
    wdOrientLandscape: 1,

    // 纸张大小 WdPaperSize
    wdPaperA4: 7,

    // 版式网格模式 WdLayoutMode
    wdLayoutModeDefault: 0,
    wdLayoutModeGrid: 1,
    wdLayoutModeLineGrid: 2,
    wdLayoutModeGenko: 3,

    // 行距规则 WdLineSpacing
    wdLineSpaceSingle: 0,
    wdLineSpace1pt5: 1,
    wdLineSpaceDouble: 2,
    wdLineSpaceAtLeast: 3,
    wdLineSpaceExactly: 4,
    wdLineSpaceMultiple: 5,

    // 视图与页眉页脚定位 WdViewType & WdSeekView
    wdNormalView: 1,
    wdOutlineView: 2,
    wdPrintView: 3,
    wdSeekMainDocument: 0,
    wdSeekCurrentPageHeader: 9,
    wdSeekCurrentPageFooter: 10,
    wdSeekPrimaryHeader: 1,
    wdSeekPrimaryFooter: 4,
    wdSeekEvenPagesHeader: 3,
    wdSeekEvenPagesFooter: 6,

    // 页眉页脚索引 WdHeaderFooterIndex
    wdHeaderFooterPrimary: 1,
    wdHeaderFooterFirstPage: 2,
    wdHeaderFooterEvenPages: 3,

    // 边框类型与样式 WdBorderType / WdLineStyle / WdLineWidth
    wdBorderTop: -1,
    wdBorderLeft: -2,
    wdBorderBottom: -3,
    wdBorderRight: -4,
    wdBorderHorizontal: -5,
    wdBorderVertical: -6,
    wdLineStyleNone: 0,
    wdLineStyleSingle: 1,
    wdLineStyleDouble: 7,
    wdLineWidth025pt: 2,
    wdLineWidth050pt: 4,
    wdLineWidth075pt: 6,
    wdLineWidth100pt: 8,
    wdLineWidth150pt: 12,
    wdLineWidth225pt: 18,
    wdLineWidth300pt: 24,

    // 颜色 WdColor
    wdColorAutomatic: -16777216,
    wdColorBlack: 0,
    wdColorRed: 255,
    wdColorBlue: 16711680,
    wdColorGreen: 32768,
    wdColorPink: 16711935,
    wdColorOrange: 26367,
    wdColorBrown: 13209,
    wdColorTeal: 8421376,
    wdColorViolet: 8388736,

    // 查找与替换 WdFindWrap / WdReplace
    wdFindStop: 0,
    wdFindAsk: 1,
    wdFindContinue: 1,
    wdReplaceNone: 0,
    wdReplaceOne: 1,
    wdReplaceAll: 2,

    // 大纲级别 WdOutlineLevel
    wdOutlineLevel1: 1,
    wdOutlineLevel2: 2,
    wdOutlineLevel3: 3,
    wdOutlineLevel4: 4,
    wdOutlineLevel5: 5,
    wdOutlineLevelBodyText: 10,

    // 列表编号样式 WdListNumberStyle / WdTrailingCharacter / WdNumberType
    wdListNumberStyleArabic: 0,
    wdListNumberStyleNumberInCircle: 18,
    wdListNumberStyleSimpChinNum3: 39,
    wdTrailingTab: 0,
    wdTrailingSpace: 1,
    wdTrailingNone: 2,
    wdNumberParagraph: 1,
    wdNumberListNum: 2,
    wdNumberAllNumbers: 3,

    // 折叠方向与移动单位 WdCollapseDirection / WdUnits
    wdCollapseEnd: 0,
    wdCollapseStart: 1,
    wdCharacter: 1,
    wdWord: 2,
    wdSentence: 3,
    wdParagraph: 4,
    wdLine: 5,
    wdStory: 6,
    wdSection: 8,
    wdPage: 9,
    wdExtend: 1,

    // 分隔符 WdBreakType
    wdSectionBreakNextPage: 2,
    wdPageBreak: 7,

    // 域类型 WdFieldType
    wdFieldEmpty: -1,
    wdFieldPage: 33,

    // 着重号 WdEmphasisMark
    wdEmphasisMarkNone: 0,
    wdEmphasisMarkOverSolidCircle: 1,
    wdEmphasisMarkOverComma: 2,
    wdEmphasisMarkOverWhiteCircle: 3,
    wdEmphasisMarkUnderSolidCircle: 4,

    // 样式类型与内置样式 WdStyleType / WdBuiltinStyle
    wdStyleTypeParagraph: 1,
    wdStyleTypeCharacter: 2,
    wdStyleNormal: -1,
    wdStyleHeading1: -2,
    wdStyleHeading2: -3,
    wdStyleHeading3: -4,
    wdStyleHeading4: -5,
    wdStyleHeading5: -6,
    wdStyleSubtitle: -75,

    // 表格对齐与定位 WdRowAlignment / WdCellVerticalAlignment / WdRelativeVerticalPosition / WdFramePosition
    wdAlignRowLeft: 0,
    wdAlignRowCenter: 1,
    wdAlignRowRight: 2,
    wdCellAlignVerticalTop: 0,
    wdCellAlignVerticalCenter: 1,
    wdCellAlignVerticalBottom: 3,
    wdRelativeVerticalPositionMargin: 0,
    wdRelativeVerticalPositionPage: 1,
    wdFrameBottom: -999997,
    wdRowHeightAuto: 0,
    wdRowHeightAtLeast: 1,
    wdRowHeightExactly: 2,

    // 图形环绕 WdWrapType & MsoZOrderCmd
    wdWrapInline: 7,
    wdWrapSquare: 0,
    wdWrapTight: 1,
    wdWrapThrough: 2,
    wdWrapNone: 3,
    wdWrapTopBottom: 4,
    wdWrapBehind: 5,
    wdWrapFront: 6,
    msoBringInFrontOfText: 4,
    msoSendBehindText: 5,

    // 保存与导出格式 WdSaveFormat / WdExportFormat
    wdFormatDocument97: 0,
    wdFormatDocument: 0,
    wdFormatText: 2,
    wdFormatHTML: 8,
    wdFormatFilteredHTML: 10,
    wdFormatXML: 11,
    wdFormatDocumentDefault: 16,
    wdFormatPDF: 17,
    wdExportFormatPDF: 17,

    // 文档保护 WdProtectionType
    wdNoProtection: -1,
    wdAllowOnlyRevisions: 0,
    wdAllowOnlyComments: 1,
    wdAllowOnlyFormFields: 2,
    wdAllowOnlyReading: 3
};

/**
 * 厘米转磅值 (1 cm = 28.3464567 pt)
 */
function cmToPt(cm) {
    return parseFloat(cm) * 28.3464567;
}

/**
 * 毫米转磅值 (1 mm = 2.83464567 pt)
 */
function mmToPt(mm) {
    return parseFloat(mm) * 2.83464567;
}

/**
 * 获取 WPS Writer Application 实例
 */
function getWpsApp() {
    if (typeof wps !== "undefined" && typeof wps.WpsApplication === "function") {
        return wps.WpsApplication();
    }
    if (typeof Application !== "undefined") {
        return Application;
    }
    if (typeof window !== "undefined" && window.wps && typeof window.wps.WpsApplication === "function") {
        return window.wps.WpsApplication();
    }
    return null;
}
