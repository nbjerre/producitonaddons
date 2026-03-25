define(["../model/models"], function (models) {

  class copyroutedialogDialogController {

    async onInit(oEnv, oEvent) {
      const param = await oEvent.getParameter("context");

      const oView = await oEnv.ActiveView();
      await oView.setCustomizedData(
        {
          targetItemCode:   param?.targetItemCode ?? "",
          sourceItemCode:   "",
          candidates:       [],
        },
        "data"
      );
    }

    async onDataLoad(oEnv, oEvent) {
      await this._loadCandidates(oEnv);
    }

    async _loadCandidates(oEnv) {
  const oView = await oEnv.ActiveView();
  await oView.showBusy();

  try {
    const service = await oEnv.getService();

    const resp = await service.ServiceLayer.get(
      `ProductTrees?$select=TreeCode,ProductDescription&$filter=TreeType eq 'iProductionTree'&$top=500`
    );

    if (!resp.isSuccess()) {
      console.error("Kunne ikke hente stykliste-kandidater");
      return;
    }

    const data = await oView.getCustomizedData("data");
    const items = resp.getData()?.value ?? [];

    data.candidates = items.map((i) => ({
      ItemCode: i.TreeCode          ?? "",
      ItemName: i.ProductDescription ?? "",
    }));

    if (data.candidates.length > 0) {
      data.sourceItemCode = data.candidates[0].ItemCode;
    }

    await oView.setCustomizedData(data, "data");

  } catch (err) {
    console.error("_loadCandidates error:", err);
  } finally {
    await oView.hideBusy();
  }
}

    async onSourceSelected(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      const data  = await oView.getCustomizedData("data");

      const grid            = await oView.Grid("candidateGrid");
      const selectedIndices = await grid.getSelectedIndices();

      if (selectedIndices && selectedIndices.length > 0) {
        data.sourceItemCode = data.candidates[selectedIndices[0]]?.ItemCode ?? "";
        await oView.setCustomizedData(data, "data");
      }
    }

    async onConfirm(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      const data  = await oView.getCustomizedData("data");

      if (!data.sourceItemCode) {
        await this._showTextMessage(oEnv, "messageTitleError", "copyRouteSelectBomToCopy", null, {}, "Error", "Select a bill of materials to copy.");
        return;
      }

      const oWindow = await oView.getWindow();
      await oWindow.close({ confirmed: true, sourceItemCode: data.sourceItemCode });
    }

    async onCancel(oEnv, oEvent) {
      const oView   = await oEnv.ActiveView();
      const oWindow = await oView.getWindow();
      await oWindow.close({ confirmed: false });
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

  return copyroutedialogDialogController;
});