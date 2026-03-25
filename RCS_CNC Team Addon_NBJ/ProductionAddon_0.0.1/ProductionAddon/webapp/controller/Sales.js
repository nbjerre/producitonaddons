
define(["./utility", "../model/models"], function (utility, models) {
  const API_BASE_URL = "https://localhost:7181";
  class Controller {
    async onInit(oEnv, oEvent) {
      this._lastSalesContext = null;
    }

    async onDataLoad(oEnv, oEvent) {
      await this._captureKnownSalesContext(oEnv, oEvent, "dataLoad");
    }

    async onExit(oEnv, oEvent) {
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

    async onButtonClick(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      await oView.showBusy();

      const service = await oEnv.getService();
      const response = await service.ServiceLayer.get("CompanyService_GetCompanyInfo");
      if (response.isSuccess()) {
        const companyInfo = response.getData();
        await this._showMessage(
          oEnv,
          "messageTitleInfo",
          await this._text(
            "salesCompanyInfoMessage",
            [companyInfo.CompanyName, companyInfo.Version],
            utility.StringJoin([
              "Company Name: " + companyInfo.CompanyName,
              "Version: " + companyInfo.Version,
            ], "\n")
          ),
          null,
          "Information"
        );
      }
      await oView.hideBusy();
    }

    async onCreateProductionForLine(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      await oView.showBusy();

      try {
        // Hent valgt linje fra grid
        const grid = await oView.Grid("wScf8XyBZ9fDKRjV61t5rM");
        const selectedIndices = await grid.getSelectedIndices();

        if (!selectedIndices || selectedIndices.length === 0) {
          await this._showTextMessage(oEnv, "messageTitleError", "salesSelectOrderLineBeforeCreateProduction", null, null, "Error", "Select an order line before creating production.");
          return;
        }

        const context = await this._resolveContext(oEnv, oEvent);

        if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
          await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderContextMissing", null, null, "Error", "Could not find the sales order (DocEntry/DocNum) in the view context.");
          return;
        }

        // Hent ordrelinjer fra ServiceLayer
        const svc = await oEnv.getService();
        const r = await svc.ServiceLayer.get(
          `Orders(${context.salesOrderDocEntry})?$select=DocEntry,DocNum,DocumentLines`
        );

        if (!r.isSuccess()) {
          await this._showTextMessage(oEnv, "messageTitleError", "salesOrderLinesFetchFailedServiceLayer", null, null, "Error", "Could not fetch order lines from Service Layer.");
          return;
        }

        const lines = r.getData().DocumentLines || [];
        const selectedLine = lines[selectedIndices[0]];

        if (!selectedLine || !selectedLine.ItemCode) {
          await this._showTextMessage(oEnv, "messageTitleError", "salesSelectedLineItemMissing", null, null, "Error", "Could not find the item number on the selected line.");
          return;
        }

        const payload = {
          salesOrderDocEntry: context.salesOrderDocEntry,
          salesOrderDocNum: context.salesOrderDocNum,
          itemCode: selectedLine.ItemCode,
          lineNum: selectedLine.LineNum,
          confirmSubBoms: false,
          subBomAdjustments: null,
        };

        let response = await this._externalRequest(oEnv, "POST", "/api/sales-production/create-for-line", payload);
        if (!response || response.success !== true) {
          await this._showMessage(oEnv, "messageTitleError", (response && response.error) || await this._text("salesUnknownApiError", null, "Unknown API error"), null, "Error");
          return;
        }

        if (response.requiresSubBomConfirmation === true && Array.isArray(response.subBoms)) {
          const confirmPayload = Object.assign({}, payload, {
            confirmSubBoms: true,
            subBomAdjustments: response.subBoms.map(function (x) {
              return {
                itemCode: x.itemCode,
                u_RCS_PQT: x.u_RCS_PQT,
                u_RCS_ONSTO: x.u_RCS_ONSTO || "Y",
              };
            }),
          });
          response = await this._externalRequest(oEnv, "POST", "/api/sales-production/create-for-line", confirmPayload);
        }

        if (response && response.success === true && response.created === true) {
          await this._showMessage(oEnv, "messageTitleOk", response.message || await this._text("salesProductionOrderCreated", null, "Production order created."), null, "OK");
        } else {
          await this._showMessage(oEnv, "messageTitleInfo", (response && response.message) || await this._text("salesNoProductionOrderCreated", null, "No production order was created."), null, "Information");
        }

      } catch (err) {
        await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
      } finally {
        await oView.hideBusy();
      }
    }

    async onCreateAllProductions(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      await oView.showBusy();

      try {
        const context = await this._resolveContext(oEnv, oEvent);

        if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
          await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderContextMissing", null, null, "Error", "Could not find the sales order (DocEntry/DocNum) in the view context.");
          return;
        }

        const payload = {
          salesOrderDocEntry: context.salesOrderDocEntry,
          salesOrderDocNum: context.salesOrderDocNum,
          confirmSubBoms: false,
          subBomAdjustments: null,
        };

        let response = await this._externalRequest(oEnv, "POST", "/api/sales-production/create-all", payload);
        if (!response || response.success !== true) {
          await this._showMessage(oEnv, "messageTitleError", (response && response.error) || await this._text("salesUnknownApiError", null, "Unknown API error"), null, "Error");
          return;
        }

        if (response.requiresSubBomConfirmation === true && Array.isArray(response.subBoms)) {
          const confirmPayload = Object.assign({}, payload, {
            confirmSubBoms: true,
            subBomAdjustments: response.subBoms.map(function (x) {
              return {
                itemCode: x.itemCode,
                u_RCS_PQT: x.u_RCS_PQT,
                u_RCS_ONSTO: x.u_RCS_ONSTO || "Y",
              };
            }),
          });
          response = await this._externalRequest(oEnv, "POST", "/api/sales-production/create-all", confirmPayload);

        }

        if (response && response.success === true && response.created === true) {
          await this._showMessage(
            oEnv,
            "messageTitleOk",
            await this._text(
              "salesCreateAllSummary",
              [response.createdCount || 0, response.failedCount || 0, response.skippedCount || 0],
              "Production orders created: {0}. Failed: {1}. Skipped: {2}."
            ),
            null,
            "OK"
          );
        } else {
          await this._showMessage(oEnv, "messageTitleInfo", (response && response.message) || await this._text("salesNoProductionOrdersCreated", null, "No production orders were created."), null, "Information");
        }
      } catch (err) {
        await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
      } finally {
        await oView.hideBusy();
      }
    }
    async onCancelAllProductions(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      await oView.showBusy();

      try {
        const context = await this._resolveContext(oEnv, oEvent);

        if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
          await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderContextMissing", null, null, "Error", "Could not find the sales order (DocEntry/DocNum) in the view context.");
          return;
        }

        const payload = {
          salesOrderDocEntry: context.salesOrderDocEntry,
          salesOrderDocNum: context.salesOrderDocNum,
        };

        const response = await this._externalRequest(oEnv, "POST", "/api/sales-production/cancel-all", payload);
        if (!response || response.success !== true) {
          await this._showMessage(oEnv, "messageTitleError", (response && response.error) || await this._text("salesUnknownApiError", null, "Unknown API error"), null, "Error");
          return;
        }

        if (response.cancelled === true) {
          await this._showMessage(
            oEnv,
            "messageTitleOk",
            await this._text("salesCancelAllSummary", [response.cancelledCount || 0, response.failedCount || 0], "Cancelled: {0}. Failed: {1}."),
            null,
            "OK"
          );
        } else {
          await this._showMessage(oEnv, "messageTitleInfo", response.message || await this._text("salesNoProductionOrdersCancelled", null, "No production orders were cancelled."), null, "Information");
        }
      } catch (err) {
        await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
      } finally {
        await oView.hideBusy();
      }
    }
    async onPlanAll(oEnv, oEvent) {
      await this._runPlanOrderAction(oEnv, oEvent, "BtnPR");
    }

    async onPlanUnlimited(oEnv, oEvent) {
      await this._runPlanOrderAction(oEnv, oEvent, "BtnPU");
    }

    async onPlanLine(oEnv, oEvent) {
      await this._runPlanOrderAction(oEnv, oEvent, "BtnPUL");
    }

    async onPlanRelease(oEnv, oEvent) {
      await this.onPlanAll(oEnv, oEvent);
    }

    async _runPlanOrderAction(oEnv, oEvent, btn) {
  const oView = await oEnv.ActiveView();

  const context = await this._resolveContext(oEnv, oEvent);

  let lineNum = null;
  if (btn === "BtnPUL") {
    lineNum = await this._resolveSelectedLineNum(oEnv, oView, context.salesOrderDocEntry);
    if (lineNum === null || lineNum === undefined) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSelectOrderLineBeforePlanLine", null, null, "Error", "Select an order line before running Plan line.");
      return;
    }
  }

  await oView.showBusy();

  try {
    if (!context.salesOrderDocEntry) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderDocEntryMissing", null, null, "Error", "Could not find the sales order DocEntry.");
      return;
    }

    const payload = { docEntry: context.salesOrderDocEntry };
    if (btn === "BtnPUL") {
      payload.lineNum = lineNum;
    }

    const pathByBtn = {
      BtnPU: "/api/plan/order/btnpu",
      BtnPUL: "/api/plan/order/btnpul",
      BtnPR: "/api/plan/order/btnpr",
    };

    const response = await this._externalRequest(oEnv, "POST", pathByBtn[btn], payload);

    if (!response || response.success !== true) {
      await this._showMessage(oEnv, "messageTitleError", (response && response.error) || (response && response.message) || await this._text("salesUnknownApiError", null, "Unknown API error"), null, "Error");
      return;
    }

    const btnMessages = {
      BtnPU: await this._text("salesPlanUnlimitedDone", null, "Plan unlimited completed."),
      BtnPUL: await this._text("salesPlanLineDone", null, "Plan line completed."),
      BtnPR: await this._text("salesPlanAllDone", null, "Plan all completed."),
    };

    await this._showMessage(oEnv, "messageTitleOk", btnMessages[btn] || await this._text("salesPlanningDone", null, "Planning completed."), null, "OK");

  } catch (err) {
    await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}

    async _resolveSelectedLineNum(oEnv, oview, docEntry) {
      const oView = await oEnv.ActiveView();
      const grid = await oView.Grid("wScf8XyBZ9fDKRjV61t5rM");
      const selectedIndices = await grid.getSelectedIndices();

      if (!selectedIndices || selectedIndices.length === 0) {
        return null;
      }

      const svc = await oEnv.getService();
      const response = await svc.ServiceLayer.get(
        `Orders(${docEntry})?$select=DocEntry,DocumentLines`
      );

      if (!response || !response.isSuccess()) {
        return null;
      }

      const lines = response.getData().DocumentLines || [];
      const selectedLine = lines[selectedIndices[0]];

      if (!selectedLine) {
        return null;
      }

      const lineNum = selectedLine.LineNum;
return (lineNum !== null && lineNum !== undefined) ? Number(lineNum) : null;
    }
    async onReleaseAllProductions(oEnv, oEvent) {
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    const context = await this._resolveContext(oEnv, oEvent);

    if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderContextMissing", null, null, "Error", "Could not find the sales order (DocEntry/DocNum) in the view context.");
      return;
    }

    const payload = {
      salesOrderDocEntry: context.salesOrderDocEntry,
      salesOrderDocNum: context.salesOrderDocNum,
    };

    const response = await this._externalRequest(oEnv, "POST", "/api/sales-production/release-all", payload);
    
    if (!response || response.success !== true) {
      await this._showMessage(oEnv, "messageTitleError", (response && response.error) || await this._text("salesUnknownApiError", null, "Unknown API error"), null, "Error");
      return;
    }

    if (response.released === true) {
      await this._showMessage(oEnv, "messageTitleOk", await this._text("salesReleaseAllSummary", [response.releasedCount || 0, response.failedCount || 0], "Released: {0}. Failed: {1}."), null, "OK");
    } else {
      await this._showMessage(oEnv, "messageTitleInfo", response.message || await this._text("salesNoProductionOrdersReleased", null, "No production orders were released."), null, "Information");
    }
  } catch (err) {
    await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}

async onFinishAllProductions(oEnv, oEvent) {
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    const context = await this._resolveContext(oEnv, oEvent);

    if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderContextMissing", null, null, "Error", "Could not find the sales order (DocEntry/DocNum) in the view context.");
      return;
    }

    const payload = {
      salesOrderDocEntry: context.salesOrderDocEntry,
      salesOrderDocNum: context.salesOrderDocNum,
    };

  const response = await this._externalRequest(oEnv, "POST", "/api/sales-production/finish-all", payload);
    if (!response || response.success !== true) {
      await this._showMessage(oEnv, "messageTitleError", (response && response.error) || await this._text("salesUnknownApiError", null, "Unknown API error"), null, "Error");
      return;
    }

    if (response.finished === true) {
      await this._showMessage(oEnv, "messageTitleOk", await this._text("salesFinishAllSummary", [response.finishedCount || 0, response.failedCount || 0], "Finished: {0}. Failed: {1}."), null, "OK");
    } else {
      await this._showMessage(oEnv, "messageTitleInfo", response.message || await this._text("salesNoProductionOrdersFinished", null, "No production orders were finished."), null, "Information");
    }
  } catch (err) {
    await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}
async onAttachCertificate(oEnv, oEvent) {
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    const grid = await oView.Grid("wScf8XyBZ9fDKRjV61t5rM");
    const selectedIndices = await grid.getSelectedIndices();

    if (!selectedIndices || selectedIndices.length === 0) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSelectOrderLineBeforeAttachCertificate", null, null, "Error", "Select an order line before attaching a certificate.");
      return;
    }

    const context = await this._resolveContext(oEnv, oEvent);

    if (!context.salesOrderDocEntry) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderDocEntryMissing", null, null, "Error", "Could not find the sales order DocEntry.");
      return;
    }

    const svc = await oEnv.getService();
    const r = await svc.ServiceLayer.get(
      `Orders(${context.salesOrderDocEntry})?$select=DocEntry,DocumentLines`
    );

    if (!r.isSuccess()) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesOrderLinesFetchFailed", null, null, "Error", "Could not fetch order lines.");
      return;
    }

    const lines = r.getData().DocumentLines || [];
    const selectedLine = lines[selectedIndices[0]];

    if (!selectedLine?.ItemCode) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSelectedLineItemMissing", null, null, "Error", "Could not find the item number on the selected line.");
      return;
    }

    const dialogContext = {
      DocEntry: context.salesOrderDocEntry,
      ItemCode: selectedLine.ItemCode,
      LineNum:  selectedLine.LineNum,
    };

   const dialog = await oEnv.newDialog({ id: "certificatdialog.dialog" });
const result = await dialog.open(dialogContext);

    if (result?.attached) {
      await this._showMessage(oEnv, "messageTitleInfo", await this._text("salesCertificatesAttachedCount", [result.count], "{0} certificate(s) attached."), null, "Information");
    }

  } catch (err) {
    console.error("onAttachCertificate error:", err);
    await this._showMessage(oEnv, "messageTitleError", err.message ?? await this._text("commonUnknownError", null, "Unknown error."), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}
async onShowCertificate(oEnv, oEvent) {
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    const grid = await oView.Grid("wScf8XyBZ9fDKRjV61t5rM");
    const selectedIndices = await grid.getSelectedIndices();

    if (!selectedIndices || selectedIndices.length === 0) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSelectOrderLineBeforeShowCertificates", null, null, "Error", "Select an order line before showing certificates.");
      return;
    }

    const context = await this._resolveContext(oEnv, oEvent);

    if (!context.salesOrderDocEntry) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderDocEntryMissing", null, null, "Error", "Could not find the sales order DocEntry.");
      return;
    }

    const svc = await oEnv.getService();
    const r = await svc.ServiceLayer.get(
      `Orders(${context.salesOrderDocEntry})?$select=DocEntry,DocumentLines`
    );

    if (!r.isSuccess()) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesOrderLinesFetchFailed", null, null, "Error", "Could not fetch order lines.");
      return;
    }

    const lines = r.getData().DocumentLines || [];
    const selectedLine = lines[selectedIndices[0]];

    if (!selectedLine?.ItemCode) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSelectedLineItemMissing", null, null, "Error", "Could not find the item number on the selected line.");
      return;
    }

    const dialogContext = {
      DocEntry: context.salesOrderDocEntry,
      ItemCode: selectedLine.ItemCode,
      LineNum:  selectedLine.LineNum,
      mode:     "view",
    };

    const dialog = await oEnv.newDialog({ id: "certificatdialog.dialog" });
    await dialog.open(dialogContext);
    await dialog.close();

  } catch (err) {
    console.error("onShowCertificate error:", err);
    await this._showMessage(oEnv, "messageTitleError", err.message ?? await this._text("commonUnknownError", null, "Unknown error."), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}
async onCopyRoute(oEnv, oEvent) {
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    // 1. Hent valgt linje fra grid
    const grid = await oView.Grid("wScf8XyBZ9fDKRjV61t5rM");
    const selectedIndices = await grid.getSelectedIndices();

    if (!selectedIndices || selectedIndices.length === 0) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesNoOrderLineSelected", null, null, "Error", "No order line selected.");
      return;
    }

    const context = await this._resolveContext(oEnv, oEvent);
    if (!context.salesOrderDocEntry) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderDocEntryMissing", null, null, "Error", "Could not find the sales order DocEntry.");
      return;
    }

    const svc = await oEnv.getService();
    const r = await svc.ServiceLayer.get(
      `Orders(${context.salesOrderDocEntry})?$select=DocEntry,DocumentLines`
    );
    if (!r.isSuccess()) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesOrderLinesFetchFailed", null, null, "Error", "Could not fetch order lines.");
      return;
    }

    const lines = r.getData().DocumentLines || [];
    const selectedLine = lines[selectedIndices[0]];

    if (!selectedLine?.ItemCode) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesNoValidItemCodeSelected", null, null, "Error", "No valid item code on the selected line.");
      return;
    }

    const itemCode = selectedLine.ItemCode;

    // 2. Tjek om stykliste allerede eksisterer
    const checkResponse = await this._externalRequest(
      oEnv, "GET", `/api/bom/check-exists/${encodeURIComponent(itemCode)}`
    );

    if (checkResponse?.exists === true) {
      await this._showMessage(oEnv, "messageTitleInfo", await this._text("salesBomAlreadyExists", [itemCode], "{0} already exists as a bill of materials."), null, "Information");
      return;
    }

    // 3. Åbn dialog til valg af kildes-stykliste
    await oView.hideBusy();

    const dialog = await oEnv.newDialog({ id: "copyroutedialog.dialog" });
    const result = await dialog.open({ targetItemCode: itemCode });

    if (!result?.confirmed || !result?.sourceItemCode) {
      return;
    }

    await oView.showBusy();

    // 4. Udfør kopiering via API
    const copyResponse = await this._externalRequest(oEnv, "POST", "/api/bom/copy-route", {
      sourceItemCode: result.sourceItemCode,
      targetItemCode: itemCode,
    });

    if (copyResponse?.success === true) {
      await this._showMessage(oEnv, "messageTitleOk", await this._text("salesRouteCopiedToItem", [itemCode], "Bill of materials copied to {0}."), null, "OK");
    } else {
      await this._showMessage(oEnv, "messageTitleError", copyResponse?.error || await this._text("salesUnknownCopyError", null, "Unknown error while copying."), null, "Error");
    }

  } catch (err) {
    await this._showMessage(oEnv, "messageTitleError", err?.message ?? String(err), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}
async onPrintProductionOrders(oEnv, oEvent) {
  const documentId = this._createClientDocumentId();
  const pendingTab = await this._openPendingTab(
    this._buildApiUrl("/api/sales-production/print-wait/" + encodeURIComponent(documentId))
  );
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    const context = await this._resolveContext(oEnv, oEvent);

    if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderContextMissing", null, null, "Error", "Could not find the sales order (DocEntry/DocNum) in the view context.");
      return;
    }

    const query = [
      context.salesOrderDocEntry ? ("salesOrderDocEntry=" + encodeURIComponent(String(context.salesOrderDocEntry))) : null,
      context.salesOrderDocNum ? ("salesOrderDocNum=" + encodeURIComponent(String(context.salesOrderDocNum))) : null,
      documentId ? ("documentId=" + encodeURIComponent(String(documentId))) : null,
    ].filter(Boolean).join("&");

    const response = await this._externalRequest(
      oEnv,
      "GET",
      "/api/sales-production/print-production-orders" + (query ? ("?" + query) : "")
    );
    if (!response || response.success !== true) {
      await this._showMessage(oEnv, "messageTitleError", (response && response.error) || await this._text("salesUnknownApiError", null, "Unknown API error"), null, "Error");
      return;
    }

    const orders = Array.isArray(response.orders) ? response.orders : [];
    const openUrl = response.openUrl || response.downloadUrl || null;
    if (!openUrl) {
      this._closePendingTab(pendingTab);
      await this._showMessage(oEnv, "messageTitleInfo", response.message || await this._text("salesNoProductionOrdersToPrint", null, "No production orders to print."), null, "Information");
      return;
    }

    let openedAutomatically = false;
    if (pendingTab) {
      openedAutomatically = true;
    } else if (openUrl) {
      openedAutomatically = this._tryOpenUrl(openUrl);
    }

    if (!openedAutomatically && openUrl) {
      openedAutomatically = this._redirectToUrl(openUrl);
    }

    const docs = orders
      .map(function (x) { return x.documentNumber; })
      .filter(function (n) { return n !== null && n !== undefined && n !== 0; })
      .slice(0, 10);

    let message = response.message || await this._text("salesPrintReadySummary", [orders.length], "Ready to print: {0} production orders.");
    if (docs.length > 0) {
      message += "\n" + await this._text("salesPrintDocNumbers", [docs.join(", ")], "Doc no.: {0}");
    }
    if (response.includedMeasureReport) {
      message += "\n" + await this._text("salesPrintMeasureReportIncluded", null, "Measurement report included where possible.");
    }
    if (response.attachmentWarningsCount > 0) {
      message += "\n" + await this._text("salesPrintAttachmentWarnings", [response.attachmentWarningsCount], "Missing attachments skipped: {0}");
    }
    if (response.renderMode) {
      message += "\n" + await this._text("salesPrintRenderMode", [response.renderMode], "Render mode: {0}");
    }

    if (openedAutomatically && response.failedCount === 0 && response.attachmentWarningsCount === 0) {
      return;
    }

    if (!openedAutomatically && openUrl) {
      const dialogMessage = [
        response.message || await this._text("salesPrintReadySummary", [orders.length], "Ready to print: {0} production orders."),
        docs.length > 0 ? await this._text("salesPrintDocNumbers", [docs.join(", ")], "Doc no.: {0}") : null,
        response.renderMode ? await this._text("salesPrintRenderMode", [response.renderMode], "Render mode: {0}") : null,
      ].filter(Boolean).join("\n");

      const dialog = await oEnv.newDialog({ id: "printpdfdialog.dialog" });
      await dialog.open({
        message: dialogMessage,
        openUrl,
      });
      await dialog.close();
      return;
    }

    await this._showMessage(oEnv, "messageTitlePrintProductionOrders", message, null, "Print production orders");
  } catch (err) {
    this._closePendingTab(pendingTab);
    await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}
async onPurchase(oEnv, oEvent) {
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    const context = await this._resolveContext(oEnv, oEvent);

    if (!context.salesOrderDocEntry) {
      await this._showTextMessage(oEnv, "messageTitleError", "salesSalesOrderDocEntryMissing", null, null, "Error", "Could not find the sales order DocEntry.");
      return;
    }

    const dialog = await oEnv.newDialog({ id: "purchasedialog.dialog" });
    await dialog.open({ DocEntry: context.salesOrderDocEntry });
    await dialog.close();

  } catch (err) {
    console.error("onPurchase error:", err);
    await this._showMessage(oEnv, "messageTitleError", err.message ?? await this._text("commonUnknownError", null, "Unknown error."), null, "Error");
  } finally {
    await oView.hideBusy();
  }
}
    async _resolveContext(oEnv, oEvent) {
      const eventParams = (oEvent && oEvent.mParameters) || {};

      const oView = await oEnv.ActiveView();
      const viewSources = await this._collectViewSources(oView);
      const context = {
        salesOrderDocEntry: null,
        salesOrderDocNum: null,
        lineNum: null,
        itemCode: null,
        contextSource: "unknown",
      };

      this._mergeContext(context, this._extractContextCandidate(eventParams, "event"));
      this._mergeContext(context, this._extractContextCandidate(oEvent, "eventObject"));

      if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
        const sEventName = eventParams.sEventName || "";
        const idFromEvent = this._parseOrdrFromString(sEventName);
        if (idFromEvent > 0) {
          this._mergeContext(context, {
            salesOrderDocEntry: idFromEvent,
            salesOrderDocNum: null,
            lineNum: null,
            itemCode: null,
            contextSource: "eventName",
          });
        }
      }

      if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
        const objectKeyContext = await this._resolveContextFromObjectKey(oView);
        this._mergeContext(context, objectKeyContext);
      }

      if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
        const knownMethodContext = await this._resolveContextFromKnownViewMethods(oView, "view:knownMethods");
        this._mergeContext(context, knownMethodContext);
      }

      if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
        for (const source of viewSources) {
          this._mergeContext(context, source.context);
          if (context.salesOrderDocEntry || context.salesOrderDocNum) {
            break;
          }
        }
      }

      if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
        const gridContext = await this._resolveContextFromGrid(oView);
        this._mergeContext(context, gridContext);
      }

      if (!context.salesOrderDocEntry && !context.salesOrderDocNum) {
        const browserContext = this._resolveContextFromBrowser();
        this._mergeContext(context, browserContext);
      }

      if (!context.salesOrderDocEntry && !context.salesOrderDocNum && this._lastSalesContext) {
        this._mergeContext(context, Object.assign({}, this._lastSalesContext, { contextSource: "cache" }));
      }

      if (!context.salesOrderDocEntry && context.salesOrderDocNum) {
        const resolvedDocEntry = await this._resolveDocEntryFromDocNum(oEnv, context.salesOrderDocNum);
        if (resolvedDocEntry) {
          context.salesOrderDocEntry = resolvedDocEntry;
          if (context.contextSource === "unknown") {
            context.contextSource = "docNumLookup";
          } else {
            context.contextSource += "+docNumLookup";
          }
        }
      }

      if (!context.salesOrderDocNum && context.salesOrderDocEntry) {
        const resolvedDocNum = await this._resolveDocNumFromDocEntry(oEnv, context.salesOrderDocEntry);
        if (resolvedDocNum) {
          context.salesOrderDocNum = resolvedDocNum;
          if (context.contextSource === "unknown") {
            context.contextSource = "docEntryLookup";
          } else {
            context.contextSource += "+docEntryLookup";
          }
        }
      }

      if (context.salesOrderDocEntry || context.salesOrderDocNum) {
        this._rememberContext(context);
      }
      return {
        salesOrderDocEntry: context.salesOrderDocEntry,
        salesOrderDocNum: context.salesOrderDocNum,
        lineNum: context.lineNum,
        itemCode: context.itemCode,
      };
    }

    async _captureKnownSalesContext(oEnv, oEvent, sourceLabel) {
      const context = await this._resolveContext(oEnv, oEvent);
      if (context.salesOrderDocEntry || context.salesOrderDocNum) {
        this._rememberContext(Object.assign({}, context, { contextSource: sourceLabel || "captured" }));
      }
      return context;
    }

    async _collectViewSources(oView) {
      const sources = [];
      for (const getter of ["getCurrentData", "getData", "getModelData", "getBindingData", "getContext", "getCustomizedData", "getCustomData"]) {
        if (typeof oView[getter] === "function") {
          try {
            const data = await oView[getter]();
            sources.push({
              getter,
              value: data,
              context: this._extractContextCandidate(data, `view:${getter}`),
            });
          } catch (e) {
            // ignore
          }
        }
      }
      return sources;
    }

    async _resolveContextFromObjectKey(oView) {
      if (!oView || typeof oView.getObjectKey !== "function") {
        return null;
      }

      try {
        const objectKey = await oView.getObjectKey();

        if (!objectKey || String(objectKey.object || "").toUpperCase() !== "ORDR") {
          return null;
        }

        const direct = this._extractContextCandidate(objectKey, "view:getObjectKey");
        if (direct && (direct.salesOrderDocEntry || direct.salesOrderDocNum)) {
          return direct;
        }

        const objectKeyFallback = await this._resolveContextFromKnownViewMethods(oView, "view:getObjectKeyFallback");
        if (objectKeyFallback && (objectKeyFallback.salesOrderDocEntry || objectKeyFallback.salesOrderDocNum)) {
          return objectKeyFallback;
        }

        const docEntry = this._extractKeyValueByName(objectKey, "DocEntry");
        const docNum = this._extractKeyValueByName(objectKey, "DocNum");
        return {
          salesOrderDocEntry: this._toNumber(docEntry),
          salesOrderDocNum: this._toNumber(docNum),
          lineNum: null,
          itemCode: null,
          contextSource: "view:getObjectKey",
        };
      } catch (e) {
        return null;
      }
    }

    async _resolveContextFromKnownViewMethods(oView, sourcePrefix) {
      if (!oView) {
        return null;
      }

      const methodNames = [
        "getCurrentData",
        "getData",
        "getModelData",
        "getBindingData",
        "getContext",
        "getCustomizedData",
        "getCustomData",
        "getViewData",
        "getRouteParameters",
        "getRouteData",
        "getParams",
        "getParameters",
        "getObject",
        "getObjectData",
        "getObjectInfo",
        "getState",
        "getSelection",
      ];

      for (const methodName of methodNames) {
        if (typeof oView[methodName] !== "function") {
          continue;
        }

        try {
          const value = await oView[methodName]();
          const candidate = this._extractContextCandidate(value, `${sourcePrefix || "view"}:${methodName}`);
          if (candidate && (candidate.salesOrderDocEntry || candidate.salesOrderDocNum)) {
            return candidate;
          }
        } catch (e) {
          // ignore
        }
      }

      return null;
    }

    async _resolveContextFromGrid(oView) {
      try {
        const grid = await oView.Grid("wScf8XyBZ9fDKRjV61t5rM");
        if (!grid) {
          return null;
        }

        const candidates = [];
        for (const getter of ["getSelectedRows", "getSelectedContexts", "getSelectedItems", "getData", "getModelData", "getBindingData", "getContext", "getRows"]) {
          if (typeof grid[getter] === "function") {
            try {
              const value = await grid[getter]();
              candidates.push(this._extractContextCandidate(value, `grid:${getter}`));
            } catch (e) {
              // ignore
            }
          }
        }

        const selectedIndices = typeof grid.getSelectedIndices === "function"
          ? await grid.getSelectedIndices()
          : [];

        if (Array.isArray(selectedIndices) && selectedIndices.length > 0) {
          const selectedIndex = Number(selectedIndices[0]);
          for (const getter of ["getRows", "getData", "getModelData", "getBindingData"]) {
            if (typeof grid[getter] === "function") {
              try {
                const value = await grid[getter]();
                const row = Array.isArray(value) ? value[selectedIndex] : value;
                candidates.push(this._extractContextCandidate(row, `grid:${getter}[${selectedIndex}]`));
              } catch (e) {
                // ignore
              }
            }
          }
        }

        for (const candidate of candidates) {
          if (candidate && (candidate.salesOrderDocEntry || candidate.salesOrderDocNum)) {
            return candidate;
          }
        }
      } catch (e) {
        // ignore
      }

      return null;
    }

    _resolveContextFromBrowser() {
      try {
        const sources = [];

        if (typeof window !== "undefined" && window && window.location) {
          sources.push({ label: "browser:href", value: window.location.href || "" });
          sources.push({ label: "browser:hash", value: window.location.hash || "" });
          sources.push({ label: "browser:search", value: window.location.search || "" });
        }

        if (typeof document !== "undefined" && document) {
          sources.push({ label: "browser:referrer", value: document.referrer || "" });
        }

        if (typeof history !== "undefined" && history) {
          sources.push({ label: "browser:historyState", value: history.state || null });
        }

        for (const source of sources) {
          const candidate = typeof source.value === "string"
            ? this._extractContextFromText(source.value, source.label)
            : this._extractContextCandidate(source.value, source.label);
          if (candidate && (candidate.salesOrderDocEntry || candidate.salesOrderDocNum)) {
            return candidate;
          }
        }
      } catch (e) {
      }

      return null;
    }

    async _resolveDocEntryFromDocNum(oEnv, docNum) {
      try {
        const svc = await oEnv.getService();
        const response = await svc.ServiceLayer.get(`Orders?$select=DocEntry,DocNum&$filter=DocNum eq ${Number(docNum)}&$top=1`);
        if (response && response.isSuccess()) {
          const data = response.getData();
          const row = Array.isArray(data && data.value) ? data.value[0] : null;
          return this._toNumber(row && row.DocEntry);
        }
      } catch (e) {
      }
      return null;
    }

    async _resolveDocNumFromDocEntry(oEnv, docEntry) {
      try {
        const svc = await oEnv.getService();
        const response = await svc.ServiceLayer.get(`Orders(${Number(docEntry)})?$select=DocEntry,DocNum`);
        if (response && response.isSuccess()) {
          const data = response.getData() || {};
          return this._toNumber(data.DocNum);
        }
      } catch (e) {
      }
      return null;
    }

    _extractContextCandidate(candidate, source) {
      if (!candidate || typeof candidate !== "object") {
        return null;
      }

      const direct = this._extractDirectContext(candidate, source);
      if (direct.salesOrderDocEntry || direct.salesOrderDocNum || direct.lineNum !== null || direct.itemCode) {
        return direct;
      }

      const deep = this._findContextDeep(candidate, 0, new Set());
      if (deep.salesOrderDocEntry || deep.salesOrderDocNum || deep.lineNum !== null || deep.itemCode) {
        deep.contextSource = source;
        return deep;
      }

      return null;
    }

    _extractContextFromText(text, source) {
      const value = this._toStr(text);
      if (!value) {
        return null;
      }

      const docEntry =
        this._firstMatchNumber(value, [
          /DocEntry\s*[=:]\s*(\d+)/i,
          /docEntry\s*[=:]\s*(\d+)/i,
          /ObjectKey\s*[=:]\s*(\d+)/i,
          /Orders\((\d+)\)/i,
          /ORDR\s*,\s*(\d+)/i,
        ]);

      const docNum = this._firstMatchNumber(value, [
        /DocNum\s*[=:]\s*(\d+)/i,
        /docNum\s*[=:]\s*(\d+)/i,
      ]);

      if (!docEntry && !docNum) {
        return null;
      }

      return {
        salesOrderDocEntry: docEntry,
        salesOrderDocNum: docNum,
        lineNum: null,
        itemCode: null,
        contextSource: source || "text",
      };
    }

    _extractDirectContext(data, source) {
      const docEntryFromKeys = this._extractKeyValueByName(data, "DocEntry");
      const docNumFromKeys = this._extractKeyValueByName(data, "DocNum");

      return {
        salesOrderDocEntry: this._toNumber(
          data.DocEntry || data.docEntry || data.docentry || data.salesOrderDocEntry || data.OrderEntry || data.orderEntry || data.ObjectKey || data.objectKey || docEntryFromKeys
        ),
        salesOrderDocNum: this._toNumber(
          data.DocNum || data.docNum || data.docnum || data.salesOrderDocNum || data.DocumentNumber || data.documentNumber || data.OrderNumber || data.orderNumber || docNumFromKeys
        ),
        lineNum: this._toLineNumber(
          data.LineNum || data.lineNum || data.linenum || data.RowNum || data.rowNum
        ),
        itemCode: this._toStr(data.ItemCode || data.itemCode || data.itemcode || data.Code || data.code),
        contextSource: source || "unknown",
      };
    }

    _extractKeyValueByName(data, keyName) {
      if (!data || !keyName) {
        return null;
      }

      const targetName = String(keyName).toLowerCase();
      const collections = [
        data.keys,
        data.Keys,
        data.values,
        data.Values,
        data.keyValues,
        data.KeyValues,
        data.objectKeys,
        data.ObjectKeys,
      ];

      for (const collection of collections) {
        if (!Array.isArray(collection)) {
          continue;
        }

        const item = collection.find(function (x) {
          const name = String((x && (x.name || x.Name || x.key || x.Key)) || "").toLowerCase();
          return name === targetName;
        });

        if (!item) {
          continue;
        }

        const value = item.value !== undefined ? item.value
          : item.Value !== undefined ? item.Value
          : item.val !== undefined ? item.val
          : item.Val !== undefined ? item.Val
          : null;

        if (value !== null && value !== undefined && value !== "") {
          return value;
        }
      }

      return null;
    }

    _findContextDeep(value, depth, visited) {
      if (!value || typeof value !== "object") {
        return { salesOrderDocEntry: null, salesOrderDocNum: null, lineNum: null, itemCode: null };
      }
      if (visited.has(value) || depth > 4) {
        return { salesOrderDocEntry: null, salesOrderDocNum: null, lineNum: null, itemCode: null };
      }

      visited.add(value);

      const direct = this._extractDirectContext(value, "unknown");
      if (direct.salesOrderDocEntry || direct.salesOrderDocNum || direct.lineNum !== null || direct.itemCode) {
        return direct;
      }

      const keys = Object.keys(value).slice(0, 50);
      for (const key of keys) {
        try {
          const nested = value[key];
          if (Array.isArray(nested)) {
            for (let i = 0; i < Math.min(nested.length, 10); i++) {
              const hit = this._findContextDeep(nested[i], depth + 1, visited);
              if (hit.salesOrderDocEntry || hit.salesOrderDocNum || hit.lineNum !== null || hit.itemCode) {
                return hit;
              }
            }
          } else if (nested && typeof nested === "object") {
            const hit = this._findContextDeep(nested, depth + 1, visited);
            if (hit.salesOrderDocEntry || hit.salesOrderDocNum || hit.lineNum !== null || hit.itemCode) {
              return hit;
            }
          }
        } catch (e) {
          // ignore
        }
      }

      return { salesOrderDocEntry: null, salesOrderDocNum: null, lineNum: null, itemCode: null };
    }

    _mergeContext(target, incoming) {
      if (!target || !incoming) {
        return target;
      }

      target.salesOrderDocEntry = target.salesOrderDocEntry || incoming.salesOrderDocEntry || null;
      target.salesOrderDocNum = target.salesOrderDocNum || incoming.salesOrderDocNum || null;
      target.lineNum = target.lineNum !== null && target.lineNum !== undefined ? target.lineNum : (incoming.lineNum !== undefined ? incoming.lineNum : null);
      target.itemCode = target.itemCode || incoming.itemCode || null;
      if ((incoming.salesOrderDocEntry || incoming.salesOrderDocNum) && (!target.contextSource || target.contextSource === "unknown")) {
        target.contextSource = incoming.contextSource || target.contextSource || "unknown";
      }
      return target;
    }

    _rememberContext(context) {
      this._lastSalesContext = {
        salesOrderDocEntry: context.salesOrderDocEntry || null,
        salesOrderDocNum: context.salesOrderDocNum || null,
        lineNum: context.lineNum !== undefined ? context.lineNum : null,
        itemCode: context.itemCode || null,
      };
    }

    _firstMatchNumber(text, regexes) {
      for (const regex of regexes) {
        const match = String(text).match(regex);
        if (match && match[1]) {
          const value = this._toNumber(match[1]);
          if (value) {
            return value;
          }
        }
      }
      return null;
    }

    _parseOrdrFromString(str) {
      if (!str) return 0;
      const decoded = str.replace(/%2C/gi, ",").replace(/%252C/gi, ",");
      const m = decoded.match(/ORDR\s*,\s*(\d+)/i);
      if (m) return Number(m[1]);
      return 0;
    }

    _toNumber(val) {
      if (val === null || val === undefined || val === "") return null;
      const n = Number(val);
      return Number.isFinite(n) && n > 0 ? n : null;
    }

    _toLineNumber(val) {
      if (val === null || val === undefined || val === "") return null;
      const n = Number(val);
      return Number.isFinite(n) && n >= 0 ? n : null;
    }

    _toStr(val) {
      if (val === null || val === undefined || val === "") return null;
      const s = String(val).trim();
      return s || null;
    }
    async onPingApi(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      await oView.showBusy();
      try {
        const response = await this._externalRequest(oEnv, "GET", "/api/ping");
        await this._showMessage(oEnv, "messageTitlePingOk", JSON.stringify(response), null, "Ping OK");
      } catch (err) {
        await this._showMessage(oEnv, "messageTitleError", (err && err.message) || String(err), null, "Error");
      } finally {
        await oView.hideBusy();
      }
    }
    async _externalRequest(oEnv, method, path, payload) {
      const service = await oEnv.getService();
      if (!service || !service.ExternalService) {
        throw new Error(await this._text("salesExternalServiceUnavailable", null, "ExternalService is not available in this context."));
      }

      const url = API_BASE_URL.replace(/\/$/, "") + (path.charAt(0) === "/" ? path : "/" + path);

      let result;
      if (String(method).toUpperCase() === "GET") {
        result = await service.ExternalService.get(url);
      } else {
        const body = (payload === undefined || payload === null)
          ? "{}"
          : (typeof payload === "string" ? payload : JSON.stringify(payload));
        result = await service.ExternalService.post(url, body, {
          "Content-Type": "application/json",
          "Accept": "application/json"
        });
      }

      const data = (result && result.raw && result.raw.data !== undefined) ? result.raw.data : result;
      if (typeof data === "string") {
        try {
          return JSON.parse(data);
        } catch (e) {
          return data;
        }
      }

      return data;
    }

    _tryOpenUrl(url) {
      if (!url) {
        return false;
      }

      const root = typeof window !== "undefined"
        ? window
        : (typeof globalThis !== "undefined" ? globalThis : null);

      try {
        if (root && typeof root.open === "function") {
          const popup = root.open(url, "_blank", "noopener,noreferrer");
          if (popup) {
            return true;
          }
        }
      } catch (e) {
      }

      try {
        if (typeof document !== "undefined" && document && typeof document.createElement === "function") {
          const link = document.createElement("a");
          link.href = url;
          link.target = "_blank";
          link.rel = "noopener noreferrer";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          return true;
        }
      } catch (e) {
      }

      return false;
    }

    _redirectToUrl(url) {
      if (!url) {
        return false;
      }

      try {
        if (typeof sap !== "undefined" && sap && sap.m && sap.m.URLHelper && typeof sap.m.URLHelper.redirect === "function") {
          sap.m.URLHelper.redirect(url, false);
          return true;
        }
      } catch (e) {
      }

      try {
        if (typeof window !== "undefined" && window && window.location && typeof window.location.assign === "function") {
          window.location.assign(url);
          return true;
        }
      } catch (e) {
      }

      try {
        if (typeof document !== "undefined" && document && typeof document.createElement === "function") {
          const link = document.createElement("a");
          link.href = url;
          link.target = "_self";
          link.rel = "noopener noreferrer";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          return true;
        }
      } catch (e) {
      }

      return false;
    }

    async _copyToClipboard(text) {
      if (!text) {
        return false;
      }

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

    _buildApiUrl(path) {
      return API_BASE_URL.replace(/\/$/, "") + (path.charAt(0) === "/" ? path : "/" + path);
    }

    _createClientDocumentId() {
      const root = typeof window !== "undefined"
        ? window
        : (typeof globalThis !== "undefined" ? globalThis : null);

      try {
        if (root && root.crypto && typeof root.crypto.randomUUID === "function") {
          return String(root.crypto.randomUUID()).replace(/[^A-Za-z0-9_-]/g, "");
        }
      } catch (e) {
      }

      return "print_" + Date.now() + "_" + Math.random().toString(36).slice(2, 12);
    }

    async _openPendingTab(url) {
      const root = typeof window !== "undefined"
        ? window
        : (typeof globalThis !== "undefined" ? globalThis : null);

      const pendingText = await this._text("salesGeneratingPdf", null, "Generating PDF...");

      try {
        if (root && typeof root.open === "function") {
          const popup = root.open(url || "", "_blank", "noopener,noreferrer");
          if (popup && popup.document) {
            try {
              popup.document.title = pendingText;
              popup.document.body.innerHTML = "<div style=\"font-family:Arial,sans-serif;padding:24px;\">" + pendingText + "</div>";
            } catch (e) {
            }
          }
          return popup || null;
        }
      } catch (e) {
      }

      try {
        if (typeof document !== "undefined" && document && typeof document.createElement === "function") {
          const link = document.createElement("a");
          link.href = url || "about:blank";
          link.target = "_blank";
          link.rel = "noopener noreferrer";
          link.style.display = "none";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          return { __synthetic: true };
        }
      } catch (e) {
      }

      return null;
    }

    _closePendingTab(pendingTab) {
      if (!pendingTab) {
        return;
      }

      if (pendingTab.__synthetic) {
        return;
      }

      try {
        if (!pendingTab.closed) {
          pendingTab.close();
        }
      } catch (e) {
      }
    }
  }

  return Controller;
});