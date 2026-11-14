/**
 * 文档加密保护与银河麒麟原生语音朗读模块 (移植自 Ribbon1.Security.cs & Ribbon1.Speech.cs)
 * 在银河麒麟系统上支持 Web Speech API (speechSynthesis) 与本地 speech-dispatcher / espeak-ng 双引擎朗读
 */
var GwSecuritySpeech = {
    isSpeaking: false,

    /**
     * 切换文档限制编辑状态
     */
    toggleProtectDoc: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        try {
            if (doc.ProtectionType === WpsEnum.wdNoProtection) {
                var pwd = prompt("请输入限制编辑（只读保护）密码（可留空）：", "");
                if (pwd === null) return;
                doc.Protect(WpsEnum.wdAllowOnlyReading, false, pwd);
                app.StatusBar = "已开启文档只读限制编辑保护";
            } else {
                var unpwd = prompt("当前文档受保护，请输入密码解除保护（若无密码直接点确定）：", "");
                if (unpwd === null) return;
                doc.Unprotect(unpwd);
                app.StatusBar = "已解除文档限制编辑保护";
            }
        } catch (e) {
            alert("切换文档限制编辑状态出错: " + e.message);
        }
    },

    /**
     * 设置文档打开密码与修改密码
     */
    setDocumentPasswords: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        var doc = app.ActiveDocument;
        try {
            var openPwd = prompt("请输入文档【打开密码】（打开文档时必须输入，留空则不修改）：", "");
            if (openPwd !== null && openPwd.length > 0) {
                doc.Password = openPwd;
            }
            var writePwd = prompt("请输入文档【修改密码】（无密码仅能只读打开，留空则不修改）：", "");
            if (writePwd !== null && writePwd.length > 0) {
                doc.WritePassword = writePwd;
            }
            if ((openPwd && openPwd.length > 0) || (writePwd && writePwd.length > 0)) {
                app.StatusBar = "文档密码已设置，请保存文档使其生效！";
                alert("密码已设置！请保存文档（Ctrl+S）以使密码生效。");
            }
        } catch (e) {
            alert("设置文档密码出错: " + e.message);
        }
    },

    /**
     * 清除文档打开密码
     */
    clearOpenPassword: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        if (!confirm("确定要清除本文档的【打开密码】吗？")) return;
        try {
            app.ActiveDocument.Password = "";
            app.StatusBar = "警告：文档打开密码已被清除！";
        } catch (e) {
            alert("清除打开密码出错: " + e.message);
        }
    },

    /**
     * 清除文档修改密码
     */
    clearWritePassword: function () {
        var app = getWpsApp();
        if (!app || !app.ActiveDocument) return;
        if (!confirm("确定要清除本文档的【修改密码】吗？")) return;
        try {
            app.ActiveDocument.WritePassword = "";
            app.StatusBar = "警告：文档修改密码已被清除！";
        } catch (e) {
            alert("清除修改密码出错: " + e.message);
        }
    },

    /**
     * 银河麒麟双通道中文语音朗读（Web Speech API + speech-dispatcher / espeak-ng 桥接）
     */
    toggleSpeak: function () {
        var app = getWpsApp();
        if (!app || !app.Selection) return;
        var self = this;

        if (this.isSpeaking) {
            this.stopSpeak();
            return;
        }

        var text = "";
        try {
            if (app.Selection.Text && app.Selection.Text.trim().length > 1) {
                text = app.Selection.Text.trim();
            } else {
                app.Selection.MoveEnd(WpsEnum.wdParagraph, 1);
                text = (app.Selection.Text || "").trim();
            }
        } catch (e) {}

        if (!text) {
            alert("请先选中要朗读的文字，或将光标置于有文字的段落中！");
            return;
        }

        var cfg = GwConfig.load();
        this.isSpeaking = true;
        app.StatusBar = "正在朗读公文内容（再次点击【语音朗读】可停止）...";

        // 优先使用 Web Speech API，若麒麟内置 WebView 不支持则自动转发至本地 speech-dispatcher/espeak-ng
        if (typeof window !== "undefined" && "speechSynthesis" in window) {
            try {
                window.speechSynthesis.cancel();
                var utter = new SpeechSynthesisUtterance(text);
                utter.lang = cfg.speechVoice || "zh-CN";
                utter.rate = Math.max(0.5, Math.min(2.0, (cfg.speechRate || 175) / 175.0));
                utter.onend = function () {
                    self.isSpeaking = false;
                    app.StatusBar = "公文朗读完成！";
                };
                utter.onerror = function () {
                    self.speakViaBridge(text, cfg);
                };
                window.speechSynthesis.speak(utter);
                return;
            } catch (e) {}
        }
        this.speakViaBridge(text, cfg);
    },

    speakViaBridge: function (text, cfg) {
        var self = this;
        var app = getWpsApp();
        GwConfig.callBridge("/api/tts-speak", {
            text: text,
            rate: cfg.speechRate || 175,
            voice: cfg.speechVoice || "zh"
        }, function (err) {
            self.isSpeaking = false;
            if (err && app) {
                app.StatusBar = "系统语音模块未就绪 (请确认已安装 espeak-ng 或 speech-dispatcher)";
            } else if (app) {
                app.StatusBar = "公文朗读完成！";
            }
        });
    },

    stopSpeak: function () {
        var app = getWpsApp();
        this.isSpeaking = false;
        if (typeof window !== "undefined" && "speechSynthesis" in window) {
            try { window.speechSynthesis.cancel(); } catch (e) {}
        }
        GwConfig.callBridge("/api/tts-stop", {}, function () {});
        if (app) app.StatusBar = "语音朗读已停止";
    },

    openSpeakOptionDialog: function () {
        var url = window.location.href.replace(/index\.html.*$/, "ui/speak_option.html");
        if (typeof wps !== "undefined" && wps.ShowDialog) {
            wps.ShowDialog(url, "银河麒麟语音朗读设置", 400, 300, false);
        } else {
            window.open(url, "_blank", "width=400,height=300");
        }
    }
};
