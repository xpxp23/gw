/**
 * 小恐龙公文排版助手 (银河麒麟 V10 / 海光 3350 专版)
 * 配置管理与本地系统桥接服务通信模块
 * 严格遵循《党政机关公文格式国家标准》(GB/T 9704—2012) 默认参数
 */
var GwConfig = (function () {
    var STORAGE_KEY = "xkl_gw_kylin_settings_v2";
    var BRIDGE_URL = "http://127.0.0.1:38890";

    var DEFAULT_SETTINGS = {
        // 页面边距 (单位: mm)
        topMarginMm: 37.0,
        bottomMarginMm: 35.0,
        leftMarginMm: 28.0,
        rightMarginMm: 26.0,
        footerDistanceMm: 24.7,

        // 版心网格与行距
        linesPerPage: 22,
        charsPerLine: 28,
        useGwGrid: true,
        useExactLineSpacing: false,
        lineSpacingPt: 28.0,
        embedFonts: false,
        showHiddenText: false,

        // 各层级公文字体配置 (名称、字号pt、加粗、斜体)
        fonts: {
            title:      { name: "方正小标宋简体", fallback: "方正小标宋_GBK", size: 22.0, bold: false, italic: false },
            title1:     { name: "黑体",           fallback: "Noto Sans CJK SC", size: 16.0, bold: false, italic: false },
            title2:     { name: "楷体",           fallback: "楷体_GB2312",    size: 16.0, bold: false, italic: false },
            title3:     { name: "仿宋",           fallback: "仿宋_GB2312",    size: 16.0, bold: false, italic: false },
            title4:     { name: "仿宋",           fallback: "仿宋_GB2312",    size: 16.0, bold: false, italic: false },
            normal:     { name: "仿宋",           fallback: "仿宋_GB2312",    size: 16.0, bold: false, italic: false },
            head:       { name: "黑体",           fallback: "Noto Sans CJK SC", size: 16.0, bold: false, italic: false },
            signerName: { name: "楷体",           fallback: "楷体_GB2312",    size: 16.0, bold: false, italic: false },
            banji:      { name: "仿宋",           fallback: "仿宋_GB2312",    size: 14.0, bold: false, italic: false },
            pageNumber: { name: "宋体",           fallback: "Noto Serif CJK SC", size: 14.0, bold: false, italic: false }
        },

        // 西文与数字默认字体
        numberFont: "Times New Roman",
        letterFont: "Times New Roman",

        // 导出与朗读配置
        dpi: 200,
        speechRate: 175,
        speechVoice: "zh-CN",

        // 历史脚本记录
        vbaHistory: []
    };

    function cloneDefaults() {
        return JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
    }

    function load() {
        try {
            var raw = null;
            if (typeof wps !== "undefined" && wps.PluginStorage) {
                raw = wps.PluginStorage.getItem(STORAGE_KEY);
            }
            if (!raw && typeof localStorage !== "undefined") {
                raw = localStorage.getItem(STORAGE_KEY);
            }
            if (raw) {
                var parsed = JSON.parse(raw);
                var merged = cloneDefaults();
                for (var k in parsed) {
                    if (k === "fonts" && parsed.fonts) {
                        for (var fk in parsed.fonts) {
                            merged.fonts[fk] = parsed.fonts[fk];
                        }
                    } else {
                        merged[k] = parsed[k];
                    }
                }
                return merged;
            }
        } catch (e) {
            console.warn("加载配置异常，使用默认配置:", e);
        }
        return cloneDefaults();
    }

    function save(cfg) {
        try {
            var str = JSON.stringify(cfg);
            if (typeof wps !== "undefined" && wps.PluginStorage) {
                wps.PluginStorage.setItem(STORAGE_KEY, str);
            }
            if (typeof localStorage !== "undefined") {
                localStorage.setItem(STORAGE_KEY, str);
            }
            return true;
        } catch (e) {
            console.error("保存配置失败:", e);
            return false;
        }
    }

    function reset() {
        var def = cloneDefaults();
        save(def);
        return def;
    }

    /**
     * 调用本地银河麒麟 Python 桥接守护进程 (127.0.0.1:38890)
     * 提供系统语音朗读(speech-dispatcher/espeak-ng)、PDF转高清图片、打开Peony文件管理器、独立文档批处理等能力
     */
    function callBridge(endpoint, payload, callback) {
        try {
            var xhr = new XMLHttpRequest();
            xhr.open("POST", BRIDGE_URL + endpoint, true);
            xhr.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
            xhr.timeout = 30000;
            xhr.onreadystatechange = function () {
                if (xhr.readyState === 4) {
                    if (xhr.status >= 200 && xhr.status < 300) {
                        try {
                            var resp = JSON.parse(xhr.responseText);
                            if (callback) callback(null, resp);
                        } catch (err) {
                            if (callback) callback(err, null);
                        }
                    } else {
                        if (callback) callback(new Error("Bridge HTTP " + xhr.status), null);
                    }
                }
            };
            xhr.onerror = function () {
                if (callback) callback(new Error("无法连接本地麒麟桥接服务 (127.0.0.1:38890)"), null);
            };
            xhr.send(JSON.stringify(payload || {}));
        } catch (e) {
            if (callback) callback(e, null);
        }
    }

    return {
        DEFAULT_SETTINGS: DEFAULT_SETTINGS,
        BRIDGE_URL: BRIDGE_URL,
        load: load,
        save: save,
        reset: reset,
        callBridge: callBridge
    };
})();
