define([], function () {
    "use strict";

    const fallbackTitle = "Production Addon";
    const translations = {
        da: {
            appTitle: "Production Addon",
            messageTitleInfo: "Information",
            messageTitleError: "Fejl",
            messageTitleOk: "OK",
            messageTitleWarning: "Advarsel",
            messageTitleResult: "Resultat",
            messageTitlePingOk: "Ping OK",
            messageTitlePrintProductionOrders: "Print produktionsordrer",
            commonUnknownError: "Ukendt fejl.",
            salesCompanyInfoMessage: "Firmanavn: {0}\nVersion: {1}",
            salesSelectOrderLineBeforeCreateProduction: "V\u00E6lg en ordrelinje f\u00F8r du opretter produktion.",
            salesSalesOrderContextMissing: "Kunne ikke finde salgsordre (DocEntry/DocNum) i view-konteksten.",
            salesOrderLinesFetchFailedServiceLayer: "Kunne ikke hente ordrelinjer fra Service Layer.",
            salesSelectedLineItemMissing: "Kunne ikke finde varenummer p\u00E5 den valgte linje.",
            salesUnknownApiError: "Ukendt fejl fra API.",
            salesProductionOrderCreated: "Produktionsordre er oprettet.",
            salesNoProductionOrderCreated: "Ingen produktionsordre blev oprettet.",
            salesCreateAllSummary: "Produktionsordrer oprettet: {0}. Fejlet: {1}. Sprunget over: {2}.",
            salesNoProductionOrdersCreated: "Ingen produktionsordrer blev oprettet.",
            salesCancelAllSummary: "Annulleret: {0}. Fejlet: {1}.",
            salesNoProductionOrdersCancelled: "Ingen produktionsordrer blev annulleret.",
            salesSelectOrderLineBeforePlanLine: "V\u00E6lg en ordrelinje f\u00F8r du k\u00F8rer Planl\u00E6g linje.",
            salesSalesOrderDocEntryMissing: "Kunne ikke finde salgsordre DocEntry.",
            salesPlanUnlimitedDone: "Planl\u00E6g ubegr\u00E6nset udf\u00F8rt.",
            salesPlanLineDone: "Planl\u00E6g linje udf\u00F8rt.",
            salesPlanAllDone: "Planl\u00E6g alle udf\u00F8rt.",
            salesPlanningDone: "Planl\u00E6gning udf\u00F8rt.",
            salesReleaseAllSummary: "Frigivet: {0}. Fejlet: {1}.",
            salesNoProductionOrdersReleased: "Ingen produktionsordrer blev frigivet.",
            salesFinishAllSummary: "F\u00E6rdigmeldt: {0}. Fejlet: {1}.",
            salesNoProductionOrdersFinished: "Ingen produktionsordrer blev f\u00E6rdigmeldt.",
            salesSelectOrderLineBeforeAttachCertificate: "V\u00E6lg en ordrelinje f\u00F8r du tilknytter certifikat.",
            salesCertificatesAttachedCount: "{0} certifikat(er) tilknyttet.",
            salesSelectOrderLineBeforeShowCertificates: "V\u00E6lg en ordrelinje f\u00F8r du viser certifikater.",
            salesNoOrderLineSelected: "Ingen ordrelinje valgt.",
            salesOrderLinesFetchFailed: "Kunne ikke hente ordrelinjer.",
            salesNoValidItemCodeSelected: "Ingen gyldig varekode p\u00E5 valgt linje.",
            salesBomAlreadyExists: "{0} er allerede oprettet som stykliste.",
            salesRouteCopiedToItem: "Stykliste kopieret til {0}.",
            salesUnknownCopyError: "Ukendt fejl under kopiering.",
            salesNoProductionOrdersToPrint: "Ingen produktionsordrer at printe.",
            salesPrintReadySummary: "Klar til print: {0} produktionsordrer.",
            salesPrintDocNumbers: "Dok.nr.: {0}",
            salesPrintMeasureReportIncluded: "M\u00E5lerapport er inkluderet hvor muligt.",
            salesPrintAttachmentWarnings: "Manglende attachments sprunget over: {0}",
            salesPrintRenderMode: "Render mode: {0}",
            salesExternalServiceUnavailable: "ExternalService er ikke tilg\u00E6ngelig i denne kontekst.",
            salesGeneratingPdf: "Genererer PDF...",
            certificateOrderNumberFetchFailed: "Kunne ikke hente ordrenummer.",
            certificateNoCertificatesSelected: "Ingen certifikater valgt.",
            certificateAttachPartial: "{0} tilknyttet. Fejl p\u00E5: {1}",
            certificateFetchFailed: "Kunne ikke hente certifikater.",
            certificateStatusAttached: "\u2713 Tilknyttet",
            certificateActivityFailed: "Activity 1 fejlede: {0}",
            copyRouteSelectBomToCopy: "V\u00E6lg en stykliste til kopiering.",
            purchaseNoItemsFound: "Ingen indk\u00F8bsvarer fundet p\u00E5 produktionsordrerne.",
            purchaseNoSupplier: "(Ingen leverand\u00F8r)",
            purchaseSelectAtLeastOneLine: "V\u00E6lg mindst \u00E9n linje.",
            purchaseItemsMissingSupplier: "F\u00F8lgende varer mangler leverand\u00F8r: {0}",
            purchaseCreateSummary: "Oprettet {0} indk\u00F8bsordre(r). {1} fejlede.",
            purchaseAllCreateFailed: "Alle {0} indk\u00F8bsordre(r) fejlede.",
            printPdfReadyDefault: "PDF'en er klar.",
            printPdfLinkMissing: "PDF-link mangler.",
            printPdfLinkCopied: "Link kopieret til udklipsholderen.",
            printPdfCopyLinkManual: "Kopi\u00E9r link manuelt:\n{0}"
        },
        en: {
            appTitle: "Production Addon",
            messageTitleInfo: "Information",
            messageTitleError: "Error",
            messageTitleOk: "OK",
            messageTitleWarning: "Warning",
            messageTitleResult: "Result",
            messageTitlePingOk: "Ping OK",
            messageTitlePrintProductionOrders: "Print production orders",
            commonUnknownError: "Unknown error.",
            salesCompanyInfoMessage: "Company name: {0}\nVersion: {1}",
            salesSelectOrderLineBeforeCreateProduction: "Select an order line before creating production.",
            salesSalesOrderContextMissing: "Could not find the sales order (DocEntry/DocNum) in the view context.",
            salesOrderLinesFetchFailedServiceLayer: "Could not fetch order lines from Service Layer.",
            salesSelectedLineItemMissing: "Could not find the item number on the selected line.",
            salesUnknownApiError: "Unknown API error.",
            salesProductionOrderCreated: "Production order created.",
            salesNoProductionOrderCreated: "No production order was created.",
            salesCreateAllSummary: "Production orders created: {0}. Failed: {1}. Skipped: {2}.",
            salesNoProductionOrdersCreated: "No production orders were created.",
            salesCancelAllSummary: "Cancelled: {0}. Failed: {1}.",
            salesNoProductionOrdersCancelled: "No production orders were cancelled.",
            salesSelectOrderLineBeforePlanLine: "Select an order line before running Plan line.",
            salesSalesOrderDocEntryMissing: "Could not find the sales order DocEntry.",
            salesPlanUnlimitedDone: "Plan unlimited completed.",
            salesPlanLineDone: "Plan line completed.",
            salesPlanAllDone: "Plan all completed.",
            salesPlanningDone: "Planning completed.",
            salesReleaseAllSummary: "Released: {0}. Failed: {1}.",
            salesNoProductionOrdersReleased: "No production orders were released.",
            salesFinishAllSummary: "Finished: {0}. Failed: {1}.",
            salesNoProductionOrdersFinished: "No production orders were finished.",
            salesSelectOrderLineBeforeAttachCertificate: "Select an order line before attaching a certificate.",
            salesCertificatesAttachedCount: "{0} certificate(s) attached.",
            salesSelectOrderLineBeforeShowCertificates: "Select an order line before showing certificates.",
            salesNoOrderLineSelected: "No order line selected.",
            salesOrderLinesFetchFailed: "Could not fetch order lines.",
            salesNoValidItemCodeSelected: "No valid item code on the selected line.",
            salesBomAlreadyExists: "{0} already exists as a bill of materials.",
            salesRouteCopiedToItem: "Bill of materials copied to {0}.",
            salesUnknownCopyError: "Unknown error while copying.",
            salesNoProductionOrdersToPrint: "No production orders to print.",
            salesPrintReadySummary: "Ready to print: {0} production orders.",
            salesPrintDocNumbers: "Doc no.: {0}",
            salesPrintMeasureReportIncluded: "Measurement report included where possible.",
            salesPrintAttachmentWarnings: "Missing attachments skipped: {0}",
            salesPrintRenderMode: "Render mode: {0}",
            salesExternalServiceUnavailable: "ExternalService is not available in this context.",
            salesGeneratingPdf: "Generating PDF...",
            certificateOrderNumberFetchFailed: "Could not fetch the order number.",
            certificateNoCertificatesSelected: "No certificates selected.",
            certificateAttachPartial: "{0} attached. Errors on: {1}",
            certificateFetchFailed: "Could not fetch certificates.",
            certificateStatusAttached: "\u2713 Attached",
            certificateActivityFailed: "Activity 1 failed: {0}",
            copyRouteSelectBomToCopy: "Select a bill of materials to copy.",
            purchaseNoItemsFound: "No purchase items were found on the production orders.",
            purchaseNoSupplier: "(No supplier)",
            purchaseSelectAtLeastOneLine: "Select at least one line.",
            purchaseItemsMissingSupplier: "The following items are missing a supplier: {0}",
            purchaseCreateSummary: "Created {0} purchase order(s). {1} failed.",
            purchaseAllCreateFailed: "All {0} purchase order(s) failed.",
            printPdfReadyDefault: "The PDF is ready.",
            printPdfLinkMissing: "PDF link is missing.",
            printPdfLinkCopied: "Link copied to clipboard.",
            printPdfCopyLinkManual: "Copy the link manually:\n{0}"
        },
        de: {
            appTitle: "Production Addon",
            messageTitleInfo: "Information",
            messageTitleError: "Fehler",
            messageTitleOk: "OK",
            messageTitleWarning: "Warnung",
            messageTitleResult: "Ergebnis",
            messageTitlePingOk: "Ping OK",
            messageTitlePrintProductionOrders: "Produktionsauftr\u00E4ge drucken",
            commonUnknownError: "Unbekannter Fehler.",
            salesCompanyInfoMessage: "Firmenname: {0}\nVersion: {1}",
            salesSelectOrderLineBeforeCreateProduction: "W\u00E4hlen Sie eine Auftragszeile aus, bevor Sie die Produktion erstellen.",
            salesSalesOrderContextMissing: "Der Verkaufsauftrag (DocEntry/DocNum) konnte im View-Kontext nicht gefunden werden.",
            salesOrderLinesFetchFailedServiceLayer: "Auftragszeilen konnten aus dem Service Layer nicht geladen werden.",
            salesSelectedLineItemMissing: "Die Artikelnummer der ausgew\u00E4hlten Zeile konnte nicht gefunden werden.",
            salesUnknownApiError: "Unbekannter API-Fehler.",
            salesProductionOrderCreated: "Produktionsauftrag wurde erstellt.",
            salesNoProductionOrderCreated: "Es wurde kein Produktionsauftrag erstellt.",
            salesCreateAllSummary: "Produktionsauftr\u00E4ge erstellt: {0}. Fehlgeschlagen: {1}. \u00DCbersprungen: {2}.",
            salesNoProductionOrdersCreated: "Es wurden keine Produktionsauftr\u00E4ge erstellt.",
            salesCancelAllSummary: "Storniert: {0}. Fehlgeschlagen: {1}.",
            salesNoProductionOrdersCancelled: "Es wurden keine Produktionsauftr\u00E4ge storniert.",
            salesSelectOrderLineBeforePlanLine: "W\u00E4hlen Sie eine Auftragszeile aus, bevor Sie Zeile planen ausf\u00FChren.",
            salesSalesOrderDocEntryMissing: "Der DocEntry des Verkaufsauftrags konnte nicht gefunden werden.",
            salesPlanUnlimitedDone: "Unbegrenzt planen abgeschlossen.",
            salesPlanLineDone: "Zeile planen abgeschlossen.",
            salesPlanAllDone: "Alle planen abgeschlossen.",
            salesPlanningDone: "Planung abgeschlossen.",
            salesReleaseAllSummary: "Freigegeben: {0}. Fehlgeschlagen: {1}.",
            salesNoProductionOrdersReleased: "Es wurden keine Produktionsauftr\u00E4ge freigegeben.",
            salesFinishAllSummary: "Fertiggemeldet: {0}. Fehlgeschlagen: {1}.",
            salesNoProductionOrdersFinished: "Es wurden keine Produktionsauftr\u00E4ge fertiggemeldet.",
            salesSelectOrderLineBeforeAttachCertificate: "W\u00E4hlen Sie eine Auftragszeile aus, bevor Sie ein Zertifikat zuordnen.",
            salesCertificatesAttachedCount: "{0} Zertifikat(e) zugeordnet.",
            salesSelectOrderLineBeforeShowCertificates: "W\u00E4hlen Sie eine Auftragszeile aus, bevor Sie Zertifikate anzeigen.",
            salesNoOrderLineSelected: "Keine Auftragszeile ausgew\u00E4hlt.",
            salesOrderLinesFetchFailed: "Auftragszeilen konnten nicht geladen werden.",
            salesNoValidItemCodeSelected: "Kein g\u00FCltiger Artikelcode in der ausgew\u00E4hlten Zeile.",
            salesBomAlreadyExists: "{0} existiert bereits als St\u00FCckliste.",
            salesRouteCopiedToItem: "St\u00FCckliste wurde nach {0} kopiert.",
            salesUnknownCopyError: "Unbekannter Fehler beim Kopieren.",
            salesNoProductionOrdersToPrint: "Keine Produktionsauftr\u00E4ge zum Drucken.",
            salesPrintReadySummary: "Bereit zum Drucken: {0} Produktionsauftr\u00E4ge.",
            salesPrintDocNumbers: "Belegnr.: {0}",
            salesPrintMeasureReportIncluded: "Messbericht wurde, wo m\u00F6glich, eingeschlossen.",
            salesPrintAttachmentWarnings: "Fehlende Anh\u00E4nge \u00FCbersprungen: {0}",
            salesPrintRenderMode: "Render-Modus: {0}",
            salesExternalServiceUnavailable: "ExternalService ist in diesem Kontext nicht verf\u00FCgbar.",
            salesGeneratingPdf: "PDF wird erstellt...",
            certificateOrderNumberFetchFailed: "Die Auftragsnummer konnte nicht geladen werden.",
            certificateNoCertificatesSelected: "Keine Zertifikate ausgew\u00E4hlt.",
            certificateAttachPartial: "{0} zugeordnet. Fehler bei: {1}",
            certificateFetchFailed: "Zertifikate konnten nicht geladen werden.",
            certificateStatusAttached: "\u2713 Zugeordnet",
            certificateActivityFailed: "Aktivit\u00E4t 1 fehlgeschlagen: {0}",
            copyRouteSelectBomToCopy: "W\u00E4hlen Sie eine St\u00FCckliste zum Kopieren aus.",
            purchaseNoItemsFound: "Keine Einkaufsartikel auf den Produktionsauftr\u00E4gen gefunden.",
            purchaseNoSupplier: "(Kein Lieferant)",
            purchaseSelectAtLeastOneLine: "W\u00E4hlen Sie mindestens eine Zeile aus.",
            purchaseItemsMissingSupplier: "Folgende Artikel haben keinen Lieferanten: {0}",
            purchaseCreateSummary: "{0} Bestellung(en) erstellt. {1} fehlgeschlagen.",
            purchaseAllCreateFailed: "Alle {0} Bestellung(en) sind fehlgeschlagen.",
            printPdfReadyDefault: "Die PDF ist bereit.",
            printPdfLinkMissing: "PDF-Link fehlt.",
            printPdfLinkCopied: "Link in die Zwischenablage kopiert.",
            printPdfCopyLinkManual: "Link manuell kopieren:\n{0}"
        }
    };

    function getLocale() {
        const languages = [];

        try {
            if (typeof navigator !== "undefined" && navigator) {
                if (Array.isArray(navigator.languages)) {
                    languages.push.apply(languages, navigator.languages);
                }
                if (navigator.language) {
                    languages.push(navigator.language);
                }
            }
        } catch (err) {
        }

        for (let index = 0; index < languages.length; index += 1) {
            const value = String(languages[index] || "").toLowerCase();
            if (value.indexOf("da") === 0) {
                return "da";
            }
            if (value.indexOf("de") === 0) {
                return "de";
            }
            if (value.indexOf("en") === 0) {
                return "en";
            }
        }

        return "da";
    }

    function formatText(text, args) {
        if (!Array.isArray(args) || args.length === 0) {
            return text;
        }

        return String(text).replace(/\{(\d+)\}/g, function (match, index) {
            const resolvedIndex = Number(index);
            return args[resolvedIndex] !== undefined ? String(args[resolvedIndex]) : match;
        });
    }

    async function getText(key, args, fallbackText) {
        const locale = getLocale();
        const dictionary = translations[locale] || translations.da;
        const value = dictionary[key] || translations.da[key] || fallbackText || key;
        return formatText(value, args);
    }

    async function getAppTitle() {
        return getText("appTitle", null, fallbackTitle);
    }

    return {
        title: fallbackTitle,
        getText: getText,
        getAppTitle: getAppTitle
    };
})