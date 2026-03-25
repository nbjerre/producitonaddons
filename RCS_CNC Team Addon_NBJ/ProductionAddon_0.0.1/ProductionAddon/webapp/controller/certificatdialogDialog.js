/**
 * Certificate Dialog Controller
 *
 * Mirrors the old VB.NET Certifikater module:
 *   - Loads available certificates (OCLG Activities CntctType=2, CntctSbjct=2)
 *   - Highlights already-attached certificates
 *   - Attaches selected certificates to the chosen sales order line
 *     by creating two Activity records per certificate (same as AddMultipleCertificatesToOrderLine)
 */

define(["../model/models"], function (models) {
  class certificatdialogDialogController {

    // ─── Lifecycle ────────────────────────────────────────────────────────────

    async onInit(oEnv, oEvent) {
  const param = await oEvent.getParameter("context");
  const oView = await oEnv.ActiveView();
  await oView.setCustomizedData(
    {
      DocEntry:         param?.DocEntry  ?? "",
      ItemCode:         param?.ItemCode  ?? "",
      LineNum:          param?.LineNum   ?? 0,
      mode:             param?.mode      ?? "attach",
      searchText:       "",
      certificates:     [],
      _allCertificates: [],
      _attached:        [],
    },
    "data"
  );
}

    async onDataLoad(oEnv, oEvent) {
      await this._loadAll(oEnv);
      const oView = await oEnv.ActiveView();
  const data = await oView.getCustomizedData("data");
  
  if (data.mode === "view") {
    // Skjul tilknyt-knappen og søgeknapper
    const attachBtn = await oView.Button("footer-attach-btn");
    await attachBtn.setVisible(false);
    const searchBtn = await oView.Button("footer-search-btn");
    await searchBtn.setVisible(false);
    const showAllBtn = await oView.Button("footer-show-all-btn");
    await showAllBtn.setVisible(false);
    const filterBtn = await oView.Button("footer-filter-order-btn");
    await filterBtn.setVisible(false);
  }
    }

    async onExit(oEnv, oEvent) {}

    // ─── Header actions ───────────────────────────────────────────────────────

    /** Show all available certificates (unfiltered). */
    async onShowAll(oEnv, oEvent) {
      const oView  = await oEnv.ActiveView();
      const data   = await oView.getCustomizedData("data");
      data.searchText   = "";
      data.certificates = [...data._allCertificates];
      await oView.setCustomizedData(data, "data");
    }

    /** Filter to certificates that already belong to this sales order. */
    async onFilterByOrder(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      const data  = await oView.getCustomizedData("data");

      // Resolve the DocNum for the current DocEntry
      const service = await oEnv.getService();
      const resp = await service.ServiceLayer.get(
        `Orders(${data.DocEntry})?$select=DocNum`
      );
      if (!resp.isSuccess()) {
        await this._showTextMessage(oEnv, "messageTitleError", "certificateOrderNumberFetchFailed", null, {}, "Error", "Could not fetch the order number.");
        return;
      }
      const docNum = String(resp.getData()?.DocNum ?? "");

      data.certificates = data._allCertificates.filter(
        (c) => String(c.OrderDocNum) === docNum
      );
      await oView.setCustomizedData(data, "data");
    }

    /** Free-text search across CerNo, CerName, DocEntry, OrderDocNum. */
    async onSearch(oEnv, oEvent) {
      const oView = await oEnv.ActiveView();
      const data  = await oView.getCustomizedData("data");
      const q     = (data.searchText ?? "").trim().toLowerCase();

      if (!q) {
        data.certificates = [...data._allCertificates];
      } else {
        data.certificates = data._allCertificates.filter((c) =>
          [c.CerNo, c.CerName, c.DocEntry, c.OrderDocNum].some((v) =>
            String(v ?? "").toLowerCase().includes(q)
          )
        );
      }
      await oView.setCustomizedData(data, "data");
    }

    // ─── Footer actions ───────────────────────────────────────────────────────

    /** Attach all selected table rows to the sales order line. */
    async onAttach(oEnv, oEvent) {
  const oView = await oEnv.ActiveView();
  const data  = await oView.getCustomizedData("data");

  // Brug Grid ligesom Sales.js gør det
  const grid = await oView.Grid("certGrid");
  const selectedIndices = await grid.getSelectedIndices();

  if (!selectedIndices || selectedIndices.length === 0) {
    await this._showTextMessage(oEnv, "messageTitleWarning", "certificateNoCertificatesSelected", null, {}, "Warning", "No certificates selected.");
    return;
  }

  await oView.showBusy();
  const service = await oEnv.getService();
  let successCount = 0;
  const errors = [];

  for (const index of selectedIndices) {
    const cert = data.certificates[index];
    if (!cert) continue;
    try {
      await this._attachCertificate(service, cert, data);
      successCount++;
    } catch (err) {
      console.error("Attach failed for", cert.CerNo, err);
      errors.push(cert.CerNo);
    }
  }

  await oView.hideBusy();

  if (errors.length > 0) {
    await this._showMessage(
      oEnv,
      "messageTitleWarning",
      await this._text("certificateAttachPartial", [successCount, errors.join(", ")], "{0} attached. Errors on: {1}"),
      {},
      "Warning"
    );
  }

  const oWindow = await oView.getWindow();
  await oWindow.close({ attached: successCount > 0, count: successCount });
}

    async onClose(oEnv, oEvent) {
      const oView   = await oEnv.ActiveView();
      const oWindow = await oView.getWindow();
      await oWindow.close({ attached: false, count: 0 });
    }

    // ─── Private helpers ──────────────────────────────────────────────────────

    /**
     * Load all available certificates (OCLG Activities where CntctType=2 and CntctSbjct=2)
     * and mark which ones are already attached to this line.
     */
    async _loadAll(oEnv) {
      const oView  = await oEnv.ActiveView();
      const data   = await oView.getCustomizedData("data");
      const service = await oEnv.getService();

      await oView.showBusy();
      try {
        // All certificates: ActivityType=2 (Certifikat Vare), Subject=2
        const resp = await service.ServiceLayer.get(
          `Activities?$select=ActivityCode,Phone,Details,Notes,DocEntry,DocNum,CardCode,ParentObjectId` +
          `&$filter=ActivityType eq 2 and Subject eq 2` +
          `&$orderby=ActivityCode desc&$top=500`
        );

        if (!resp.isSuccess()) {
          await this._showTextMessage(oEnv, "messageTitleError", "certificateFetchFailed", null, {}, "Error", "Could not fetch certificates.");
          return;
        }

        const rows = resp.getData()?.value ?? [];

        // Fetch already-attached ClgCodes for this line so we can highlight them
        const attachedSet = await this._loadAttachedCodes(service, data);

        const certs = rows.map(async (r) => ({
          ClgCode:      r.ActivityCode,
          CerNo:        r.Phone    ?? "",   // Tel = certificate number
          CerName:      r.Details  ?? "",   // Details = certificate name
          Notes:        r.Notes    ?? "",
          DocEntry:     r.DocEntry ?? "",   // linked item code
          OrderDocNum:  r.DocNum   ?? "",
          CardCode:     r.CardCode ?? "",
          ServiceCallNo: r.ParentObjectId ?? "",
          // Visual highlight for already-attached rows
          statusLabel:  attachedSet.has(r.ActivityCode) ? await this._text("certificateStatusAttached", null, "Attached") : "",
          _attached:    attachedSet.has(r.ActivityCode),
        }));

        data._allCertificates = certs;
// Hvis view-mode: vis kun tilknyttede
if (data.mode === "view") {
  data.certificates = certs.filter(c => c._attached);
} else {
  data.certificates = [...certs];
}
data._attached = [...attachedSet];
await oView.setCustomizedData(data, "data");
      } finally {
        await oView.hideBusy();
      }
    }

    /**
     * Returns a Set of ClgCodes already linked to the current DocEntry + ItemCode.
     * Mirrors the SQL in markAlreadyAddedCertificates / frmShowCertificates_LoadMatrix1.
     *
     * Logic: find Activities (CntctType=1, Subject=4) that are linked to DocEntry (the order)
     *        and have a child activity (prevActivity) linked to ItemCode — these are "already attached".
     *
     * We approximate this via two SL calls:
     *  1. Find parent activities on the order (DocType=17, DocEntry=orderDocEntry, Subject=4, ActivityType=1)
     *  2. Find child activities linked to those parents (PreviousActivity in [...]) with DocType=4, DocEntry=itemCode
     */
    async _loadAttachedCodes(service, data) {
  const attached = new Set();
  try {
    // Find alle ordre-links på denne ordre
    const parentResp = await service.ServiceLayer.get(
      `Activities?$select=ActivityCode,PreviousActivity` +
      `&$filter=ActivityType eq 1 and Subject eq 3` +
      ` and DocType eq '17' and DocEntry eq '${data.DocEntry}'` +
      `&$top=200`
    );
    if (!parentResp.isSuccess()) return attached;
    const parents = parentResp.getData()?.value ?? [];
    if (parents.length === 0) return attached;

    // Hent de certifikat ClgCodes der er linket til ordren
    const certClgCodes = new Set(parents.map((p) => p.PreviousActivity));

    // For hvert certifikat — tjek om der også er et vare-link til vores ItemCode
    const filter = [...certClgCodes].map((c) => `PreviousActivity eq ${c}`).join(" or ");
    const childResp = await service.ServiceLayer.get(
      `Activities?$select=ActivityCode,PreviousActivity` +
      `&$filter=ActivityType eq -1 and Subject eq 1` +
      ` and DocType eq '4' and DocEntry eq '${data.ItemCode}'` +
      ` and (${filter})` +
      `&$top=200`
    );
    if (!childResp.isSuccess()) return attached;

    for (const c of childResp.getData()?.value ?? []) {
      attached.add(c.PreviousActivity); // certifikat ClgCode
    }

  } catch (err) {
    console.error("_loadAttachedCodes error", err);
  }
  return attached;
}
    /**
     * Creates the two Activity records that link a certificate to a sales order line.
     * Mirrors AddSingleCertificateToOrderLine / AddMultipleCertificatesToOrderLine.
     *
     *  Record 1 → linked to the Sales Order (DocType=17)
     *    ActivityType = 1   (Certifikat Ordre)
     *    Subject      = 4   (Cert. Ordre Nr)
     *
     *  Record 2 → linked to the Item (DocType=4)
     *    ActivityType = 1   (Certifikat Ordre)
     *    Subject      = 1   (Certifikat Vare)
     */
    async _attachCertificate(service, cert, data) {
      // ── Activity 1: link to Sales Order ──────────────────────────────────
      const act1 = {
  CardCode:         cert.CardCode,
  DocType:          '17',
  DocEntry:         String(data.DocEntry),
  PreviousActivity: cert.ClgCode,
  ActivityType:     1,
  Subject:          3,  // ← var 4, nu 3
  Phone:            cert.CerNo,
  Details:          cert.CerName,
  Notes:            cert.Notes,
};

      const resp1 = await service.ServiceLayer.post("Activities", act1);
      if (!resp1.isSuccess()) {
        throw new Error(await this._text("certificateActivityFailed", [resp1.getErrorMessage?.() ?? "unknown"], "Activity 1 failed: {0}"));
      }
      const newClgCode = resp1.getData()?.ActivityCode;

      // ── Activity 2: link to Item ─────────────────────────────────────────
      const act2 = {
  CardCode:         cert.CardCode,
  DocType:          '4',
  DocEntry:         String(data.ItemCode),
  PreviousActivity: cert.ClgCode,
  ActivityType:     -1,
  Subject:          1,
  Phone:            cert.CerNo,
  Details:          cert.CerName,
  Notes:            cert.Notes,
};

      const resp2 = await service.ServiceLayer.post("Activities", act2);
      if (!resp2.isSuccess()) {
        // Not fatal — the order link (act1) is the critical one
      }
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

  return certificatdialogDialogController;
});