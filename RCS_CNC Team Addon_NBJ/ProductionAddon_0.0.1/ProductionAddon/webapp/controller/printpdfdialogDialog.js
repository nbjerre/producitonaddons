define(["../model/models"], function (models) {

  class printpdfdialogDialogController {

    async onInit(oEnv, oEvent) {
      const param = await oEvent.getParameter("context");
      const oView = await oEnv.ActiveView();
      const oWindow = await oView.getWindow();

      await this._setWindowTitle(oWindow);

      await oView.setCustomizedData(
        {
          message: param?.message ?? await this._text("printPdfReadyDefault", null, "The PDF is ready."),
          openUrl: param?.openUrl ?? "",
        },
        "data"
      );
    }

    async onCopyLink(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      const data = await oView.getCustomizedData("data");
      const url = data?.openUrl || "";

      if (!url) {
        await this._showTextMessage(oEnv, "messageTitleError", "printPdfLinkMissing", null, {}, "Error", "PDF link is missing.");
        return;
      }

      const copied = await this._copyToClipboard(url);
      await this._showMessage(
        oEnv,
        "messageTitleInfo",
        copied
          ? await this._text("printPdfLinkCopied", null, "Link copied to clipboard.")
          : await this._text("printPdfCopyLinkManual", [url], "Copy the link manually:\n{0}"),
        {},
        "Information"
      );

      if (copied) {
        const oWindow = await oView.getWindow();
        await oWindow.close();
      }
    }

    async onClose(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      const oWindow = await oView.getWindow();
      await oWindow.close();
    }

    async _text(key, args, fallbackText) {
      return models.getText(key, args, fallbackText);
    }

    async _showMessage(oEnv, titleKey, message, options, fallbackTitle) {
      const mergedOptions = Object.assign({ title: await models.getAppTitle() }, options || {});
      return oEnv.showMessageBox(
        await this._text(titleKey, null, fallbackTitle || titleKey),
        message,
        mergedOptions
      );
    }

    async _showTextMessage(oEnv, titleKey, messageKey, args, options, fallbackTitle, fallbackMessage) {
      return this._showMessage(
        oEnv,
        titleKey,
        await this._text(messageKey, args, fallbackMessage),
        options,
        fallbackTitle
      );
    }

    async _setWindowTitle(oWindow) {
      if (!oWindow) {
        return;
      }

      const title = await this._text("printPdfDialogTitle", null, "Print production orders");

      try {
        if (typeof oWindow.setTitle === "function") {
          await oWindow.setTitle(title);
          return;
        }
      } catch (e) {
      }

      try {
        if (typeof oWindow.setCaption === "function") {
          await oWindow.setCaption(title);
        }
      } catch (e) {
      }
    }

    async _copyToClipboard(text) {
      try {
        if (typeof navigator !== "undefined" && navigator && navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
          await navigator.clipboard.writeText(text);
          return true;
        }
      } catch (e) {
      }

      try {
        if (typeof document !== "undefined" && document && typeof document.createElement === "function") {
          const textarea = document.createElement("textarea");
          textarea.value = text;
          textarea.setAttribute("readonly", "readonly");
          textarea.style.position = "fixed";
          textarea.style.opacity = "0";
          document.body.appendChild(textarea);
          textarea.select();
          textarea.setSelectionRange(0, textarea.value.length);
          const copied = typeof document.execCommand === "function" ? document.execCommand("copy") : false;
          document.body.removeChild(textarea);
          return !!copied;
        }
      } catch (e) {
      }

      return false;
    }
  }

  return printpdfdialogDialogController;
});
