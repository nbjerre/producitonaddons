define(["../model/models"], function (models) {

  class purchasedialogDialogController {

    async onInit(oEnv, oEvent) {
      const param = await oEvent.getParameter("context");

      const oView = await oEnv.ActiveView();
      await oView.setCustomizedData(
        {
          DocEntry:  param?.DocEntry ?? "",
          lines:     [],
          _allLines: [],
        },
        "data"
      );
    }

    async onDataLoad(oEnv, oEvent) {
      await this._loadLines(oEnv);
    }

    // ── Hent produktionsordrelinjer ──────────────────────────────────────────
    async _loadLines(oEnv) {
      const oView = await oEnv.ActiveView();
      const data  = await oView.getCustomizedData("data");
      await oView.showBusy();

      try {
        const service = await oEnv.getService();

        const prodResp = await service.ServiceLayer.get(
          `ProductionOrders?$select=AbsoluteEntry,ItemNo,Project,ProductionOrderLines` +
          `&$filter=ProductionOrderOriginEntry eq ${data.DocEntry}` +
          `&$top=200`
        );

        if (!prodResp.isSuccess()) {
          console.error("Kunne ikke hente produktionsordrer");
          return;
        }

        const prodOrders = prodResp.getData()?.value ?? [];
        const rawLines = [];

        for (const po of prodOrders) {
          for (const line of po.ProductionOrderLines ?? []) {
            // Fjern kommentar når CNCT_ varer er i systemet:
            // if (!line.ItemNo || !line.ItemNo.startsWith("CNCT_")) continue;
            rawLines.push({
              ProdDocEntry: po.AbsoluteEntry,
              Project:      po.Project ?? "",
              ItemCode:     line.ItemNo,
              PlannedQty:   line.PlannedQuantity ?? 0,
              StartDate:    line.StartDate ? line.StartDate.substring(0, 10) : "",
              EndDate:      line.EndDate   ? line.EndDate.substring(0, 10)   : "",
              LineNum:      line.LineNumber,
            });
          }
        }

        if (rawLines.length === 0) {
          data.lines     = [];
          data._allLines = [];
          await oView.setCustomizedData(data, "data");
          await this._showTextMessage(oEnv, "messageTitleInfo", "purchaseNoItemsFound", null, {}, "Information", "No purchase items were found on the production orders.");
          return;
        }

        // Hent leverandørinfo for unikke varenumre
        const uniqueItems = [...new Set(rawLines.map((l) => l.ItemCode))];
        const itemFilter  = uniqueItems.map((c) => `ItemCode eq '${c}'`).join(" or ");

        const itemResp = await service.ServiceLayer.get(
          `Items?$select=ItemCode,ItemName,Mainsupplier&$filter=${itemFilter}`
        );

        const itemMap = {};
        if (itemResp.isSuccess()) {
          for (const item of itemResp.getData()?.value ?? []) {
            itemMap[item.ItemCode] = {
              ItemName:     item.ItemName     ?? "",
              SupplierCode: item.Mainsupplier ?? "",
            };
          }
        }

        // Hent leverandørnavne
        const supplierCodes = [
          ...new Set(Object.values(itemMap).map((i) => i.SupplierCode).filter(Boolean)),
        ];

        const supplierMap = {};
        if (supplierCodes.length > 0) {
          const supFilter = supplierCodes.map((c) => `CardCode eq '${c}'`).join(" or ");
          const supResp   = await service.ServiceLayer.get(
            `BusinessPartners?$select=CardCode,CardName&$filter=${supFilter}`
          );
          if (supResp.isSuccess()) {
            for (const bp of supResp.getData()?.value ?? []) {
              supplierMap[bp.CardCode] = bp.CardName;
            }
          }
        }

        // Byg endelig liste
        const lines = rawLines.map(async (l) => {
          const itemInfo     = itemMap[l.ItemCode] ?? {};
          const supplierCode = itemInfo.SupplierCode ?? "";
          return {
            ...l,
            ItemName:     itemInfo.ItemName ?? "",
            SupplierCode: supplierCode,
            SupplierName: supplierCode
              ? `${supplierMap[supplierCode] ?? ""} (${supplierCode})`
              : await this._text("purchaseNoSupplier", null, "(No supplier)"),
          };
        });

        data.lines     = lines;
        data._allLines = lines;
        await oView.setCustomizedData(data, "data");

      } catch (err) {
        console.error("_loadLines error:", err);
      } finally {
        await oView.hideBusy();
      }
    }

    // ── Opret indkøbsordre(r) ────────────────────────────────────────────────
    async onCreatePurchaseOrders(oEnv, oEvent) {
      const oView          = await oEnv.ActiveView();
      const grid           = await oView.Grid("purchGrid");
      const selectedIndices = await grid.getSelectedIndices();

      if (!selectedIndices || selectedIndices.length === 0) {
        await this._showTextMessage(oEnv, "messageTitleError", "purchaseSelectAtLeastOneLine", null, {}, "Error", "Select at least one line.");
        return;
      }

      const data          = await oView.getCustomizedData("data");
      const selectedLines = selectedIndices.map((i) => data.lines[i]);

      // Valider leverandør
      const missingSupplier = selectedLines.filter((l) => !l.SupplierCode);
      if (missingSupplier.length > 0) {
        const items = missingSupplier.map((l) => l.ItemCode).join(", ");
        await this._showMessage(oEnv, "messageTitleError", await this._text("purchaseItemsMissingSupplier", [items], "The following items are missing a supplier: {0}"), {}, "Error");
        return;
      }

      await oView.showBusy();
      const service = await oEnv.getService();
      const created = [];
      const failed  = [];

      try {
        // Grupper per leverandør
        const bySupplier = {};
        for (const line of selectedLines) {
          if (!bySupplier[line.SupplierCode]) bySupplier[line.SupplierCode] = [];
          bySupplier[line.SupplierCode].push(line);
        }

        for (const [supplierCode, lines] of Object.entries(bySupplier)) {
          // Tidligste slutdato som leveringsdato
          const dueDate = lines
            .map((l) => l.EndDate)
            .filter(Boolean)
            .sort()[0] ?? new Date().toISOString().substring(0, 10);

          const documentLines = lines.map((line) => ({
            ItemCode:          line.ItemCode,
            Quantity:          line.PlannedQty,
            ProjectCode:       line.Project ?? "",
            U_PPSOne_WOOrigin: line.ProdDocEntry,  // produktionsordre reference
          }));

          const payload = {
            CardCode:      supplierCode,
            DocDueDate:    dueDate,
            DocumentLines: documentLines,
          };

          const result = await service.ServiceLayer.post("PurchaseOrders", payload);

          if (!result.isSuccess()) {
            console.error("PO fejlede for", supplierCode, result);
            failed.push({ supplierCode, error: "Oprettelse fejlede" });
            continue;
          }

          created.push({ supplierCode, docEntry: result.getData()?.DocEntry });
        }

        const msg = created.length > 0
          ? await this._text("purchaseCreateSummary", [created.length, failed.length], "Created {0} purchase order(s). {1} failed.")
          : await this._text("purchaseAllCreateFailed", [failed.length], "All {0} purchase order(s) failed.");

        await this._showMessage(oEnv, "messageTitleResult", msg, {}, "Result");

      } catch (err) {
        console.error("onCreatePurchaseOrders error:", err);
        await this._showMessage(oEnv, "messageTitleError", err.message ?? await this._text("commonUnknownError", null, "Unknown error."), {}, "Error");
      } finally {
        await oView.hideBusy();
      }
    }

    // ── Luk dialog ───────────────────────────────────────────────────────────
    async onClose(oEnv, oEvent) {
      const oView   = await oEnv.ActiveView();
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
  }

  return purchasedialogDialogController;
});