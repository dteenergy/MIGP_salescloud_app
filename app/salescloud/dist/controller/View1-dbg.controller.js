sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/BindingMode",
    "sap/m/MessageBox",
    "sap/m/MessageToast"
], function (Controller, JSONModel, _BindingMode, MessageBox, MessageToast) {
    "use strict";

    return Controller.extend("com.sap.salescloud.controller.View1", {
        // var bUseMock = true; // ← CHANGE TO false WHEN CPI IS READY
        // ─────────────────────────────────────────────────────────────────────
        // onInit
        // ─────────────────────────────────────────────────────────────────────
        onInit: function () {
            //UserRole
            // this.getUserDetails();
            var oSCDetail = new JSONModel({
                Oppid: '',
                quoteId: '',        // ✅ renamed from contractid
                objectStatus: '',        // ✅ new field — 'OPP' or 'QUOTE'
                isQuoteContext: false,      // ✅ new flag — true when Quote tab in C4C
                businessPartnerId: '',        // ✅ new field — from C4C URL param
                oppCategory: '',
                status: '',
                annualGross: '',        // FDD #1
                percentUsage: '85',      // FDD #2 default 85%
                sizeofOpp: '',        // FDD #3
                // annualMWh      : '',        // FDD #4
                annualSubs: '',        // FDD #5
                priceStructure: '',        // FDD #6
                salesType: '',        // FDD #7
                contLifetimeval: '',        // FDD #8
                oppSubs: '',        // FDD #9
                term: '',        // FDD #10
                commencementLetterSent: false,  // FDD #11 (LCVP)
                commencementLetterSigned: false,  // FDD #12 (LCVP)
                // FDD #13: free-text, comma-separated values from CR&B
                // e.g. "MIGP-FIXED" | "MIGP-FLEX" | "LCVP" | "MIGP-FLEX,LCVP"
                enrolled: '',
                meteredCon: '',        // FDD #14 (LCVP)
                unmeteredCon: '',        // FDD #15 (LCVP)
                nte: '',        // FDD #16
                includeInPipeline: false,     // FDD #18
                salesPhase: '',   // ✅ ADD
                enrollmentReferralCode: '',   // SMB — Referral Code (e.g. TONY)
                enrollmentEmailId: '',        // SMB — Enrollment Email ID
                lastFetchedAt: '',   // ← persisted timestamp from DB
                lastFetchedAtISO: '',    // ← ADD: raw ISO for save
                OppIdSuggestions: [],
                // Summary bar (LCVP, shown below tablelarge when 2+ rows after Calculate)
                summarySubFee: '',
                summarySubPct: '',
                summaryMigpPct: '',
                summaryRecPct: ''
            });
            oSCDetail.setDefaultBindingMode("TwoWay");
            this.getView().setModel(oSCDetail, "oSCDetail");

            var oTabModel = new JSONModel({
                InvolvedParties: [],
                ProviderContracts: [],
                ConsumptionDetails: [],
                Prospects: [],
                lastFetchTime: "",        // ← ADD
                lastFetchVisible: false,  // ← ADD
                // lastValidateTime: "",     // ← ADD
                // lastValidateVisible: false // ← ADD
                lastFetchedAt: "",
            });
            oTabModel.setDefaultBindingMode("TwoWay");
            this.getOwnerComponent().setModel(oTabModel, "oOpportunityjmodel");

            var oRouter = sap.ui.core.UIComponent.getRouterFor(this);
            oRouter.getRoute("RouteView1").attachPatternMatched(this._onRouteMatched, this);

            // ── Show suggestions on focus (attachBrowserEvent — focus attr not supported in UI5) ──
            var oInput = this.getView().byId("il8");
            if (oInput) {
                oInput.attachBrowserEvent("focus", function () {
                    var oSCDetailModel = this.getView().getModel("oSCDetail");
                    var aAll = oSCDetailModel.getProperty("/OppIdSuggestions") || [];
                    if (aAll.length === 0) { return; }
                    oInput.destroySuggestionItems();
                    aAll.slice(0, 20).forEach(function (r) {
                        oInput.addSuggestionItem(new sap.ui.core.Item({
                            key: r.Oppid,
                            text: r.Oppid +
                                (r.oppCategory ? "  |  " + r.oppCategory : "") +
                                (r.status ? "  [" + r.status + "]" : "")
                        }));
                    });
                }.bind(this));
            }
        },

        //UserRoles
        // getUserDetails: async function () {
        //     let oUserScopeJModel = this.getOwnerComponent().getModel("oUserScopeJModel");
        //     oUserScopeJModel.setData("");

        //     let omodel = this.getOwnerComponent().getModel();

        //     // ── Shared handler: applies user result to model + sets UI state ──
        //     var fnApplyUserResult = function (oOperation) {
        //         let oResults = oOperation.getBoundContext().getObject();
        //         if (!oResults) { return; }

        //         oUserScopeJModel.setData(oResults);

        //         if (oUserScopeJModel.oData.hasAdminAccess) {
        //             this.getView().byId("idcommLetterSent").setEnabled(true);
        //             this.getView().byId("idcommLetterSigned").setEnabled(true);
        //             this.getView().byId("idbtnExport").setVisible(
        //                 this.getView().getModel("oSCDetail").getProperty("/isQuoteContext")
        //             );
        //         } else if (oUserScopeJModel.oData.hasUserAccess) {
        //             this.getView().byId("idcommLetterSent").setEnabled(false);
        //             this.getView().byId("idcommLetterSigned").setEnabled(false);
        //             this.getView().byId("idbtnExport").setVisible(false);
        //         }
        //     }.bind(this);

        //     let oOperation = omodel.bindContext("/userDetails(...)");
        //     try {
        //         await oOperation.execute().then(function () {
        //             fnApplyUserResult(oOperation);
        //         }.bind(this), function (err) {
        //             sap.m.MessageBox.error(err.message + "\n Roles not found! Contact support admin!");
        //             console.error(err.message);
        //         }.bind(this));

        //     } catch (oError) {
        //         // BUG 17908: Transient 401/role errors on BTP — retry once after a short delay
        //         var sErrMsg = (oError && oError.message) ? oError.message : String(oError);
        //         if (sErrMsg.indexOf("401") !== -1 || sErrMsg.indexOf("403") !== -1 ||
        //             sErrMsg.indexOf("No role") !== -1 || sErrMsg.indexOf("not trusted") !== -1) {
        //             console.warn("userDetails: transient auth error — retrying in 2s:", sErrMsg);
        //             setTimeout(async function () {
        //                 try {
        //                     var oRetryOp = omodel.bindContext("/userDetails(...)");
        //                     await oRetryOp.execute();
        //                     fnApplyUserResult(oRetryOp);  // ← apply result from retry
        //                     console.log("userDetails: retry succeeded.");
        //                 } catch (oRetryErr) {
        //                     console.error("userDetails retry also failed:", oRetryErr.message);
        //                     sap.m.MessageBox.error(
        //                         "Unable to verify your role. Please refresh the page.\n\n" +
        //                         "If this persists, contact your administrator to confirm your role assignment in BTP.",
        //                         { title: "Authorization Error" }
        //                     );
        //                 }
        //             }.bind(this), 2000);
        //             return;
        //         }
        //         console.error("userDetails error:", oError);
        //     }
        // },

        getUserDetails: async function () {
            let oUserScopeJModel = this.getOwnerComponent().getModel("oUserScopeJModel");
            oUserScopeJModel.setData({});  // Fix: clear with empty object, not ""

            let omodel = this.getOwnerComponent().getModel();

            var fnApplyUserResult = function (oOperation) {
                // Default-deny: hide all three controls before inspecting result
                this.getView().byId("idcommLetterSent").setVisible(false);
                this.getView().byId("idcommLetterSent").setEnabled(false);
                this.getView().byId("idcommLetterSigned").setVisible(false);
                this.getView().byId("idcommLetterSigned").setEnabled(false);
                this.getView().byId("idbtnExport").setVisible(false);
                this.getView().byId("idbtnExport").setEnabled(false);

                let oResults = oOperation.getBoundContext().getObject();
                if (!oResults) {
                    console.warn("userDetails: empty result — controls hidden by default");
                    return;
                }

                oUserScopeJModel.setData(oResults);

                if (oUserScopeJModel.oData.hasAdminAccess) {
                    // Show and enable commencement checkboxes for admin
                    this.getView().byId("idcommLetterSent").setVisible(true);
                    this.getView().byId("idcommLetterSent").setEnabled(true);
                    this.getView().byId("idcommLetterSigned").setVisible(true);
                    this.getView().byId("idcommLetterSigned").setEnabled(true);
                    // Export: visible only in quote context
                    // var bIsQuote = this.getView().getModel("oSCDetail").getProperty("/isQuoteContext");
                    // this.getView().byId("idbtnExport").setVisible(!!bIsQuote);
                    // this.getView().byId("idbtnExport").setEnabled(true);
                    // Export: visible based on category + context (admin only)
                    const oSCDetail = this.getView().getModel("oSCDetail");
                    const bIsQuoteContext = oSCDetail.getProperty("/isQuoteContext");
                    const sCategory = oSCDetail.getProperty("/oppCategory");

                    const bShowExport =
                        (sCategory === "MIGP - Small Business" && !bIsQuoteContext) ||
                        ((sCategory === "MIGP - LCVP" || sCategory === "MIGP - Dedicated Array") && bIsQuoteContext);

                    this.getView().byId("idbtnExport").setVisible(bShowExport);
                    this.getView().byId("idbtnExport").setEnabled(bShowExport);
                }
                // Non-admin: controls remain hidden from default-deny above
            }.bind(this);

            let oOperation = omodel.bindContext("/userDetails(...)");
            try {
                // Fix: no two-arg .then — let all rejections fall into catch
                await oOperation.execute();
                fnApplyUserResult(oOperation);
            } catch (oError) {
                // Hide controls on any error — deny by default
                this.getView().byId("idcommLetterSent").setVisible(false);
                this.getView().byId("idcommLetterSent").setEnabled(false);
                this.getView().byId("idcommLetterSigned").setVisible(false);
                this.getView().byId("idcommLetterSigned").setEnabled(false);
                this.getView().byId("idbtnExport").setVisible(false);
                this.getView().byId("idbtnExport").setEnabled(false);

                var sErrMsg = (oError && oError.message) ? oError.message : String(oError);
                if (sErrMsg.indexOf("401") !== -1 || sErrMsg.indexOf("403") !== -1 ||
                    sErrMsg.indexOf("No role") !== -1 || sErrMsg.indexOf("not trusted") !== -1) {
                    console.warn("userDetails: transient auth error — retrying in 2s:", sErrMsg);
                    setTimeout(async function () {
                        try {
                            var oRetryOp = omodel.bindContext("/userDetails(...)");
                            await oRetryOp.execute();
                            fnApplyUserResult(oRetryOp);
                            console.log("userDetails: retry succeeded.");
                        } catch (oRetryErr) {
                            console.error("userDetails retry also failed:", oRetryErr.message);
                            sap.m.MessageBox.error(
                                "Unable to verify your role. Please refresh the page.\n\n" +
                                "If this persists, contact your administrator to confirm your role assignment in BTP.",
                                { title: "Authorization Error" }
                            );
                        }
                    }.bind(this), 2000);
                    return;
                }
                console.error("userDetails error:", oError);
                sap.m.MessageBox.error("Roles not found! Contact support admin.");
            }
        },

        // ── Sales Phase definitions per BPD ──────────────────────────────────
        // SMB:  C4C sends "Approval"    → BTP shows from Approval onwards
        // LCVP: C4C sends "Structuring" → BTP shows from Structuring onwards
        // ─────────────────────────────────────────────────────────────────────
        _oSalesPhases: {
            "MIGP - Small Business": [
                { code: "Z0", text: "Approval" },
                { code: "Z1", text: "Customer Appointment" },
                { code: "Z2", text: "Awaiting Confirmation" },
                { code: "Z3", text: "Final Review" },
                { code: "Z4", text: "Closure" }
            ],
            "MIGP - LCVP": [
                { code: "Z5", text: "Origination" },
                { code: "Z6", text: "Introduction" },
                { code: "Z7", text: "Discovery" },
                { code: "Z8", text: "Structuring" },
                { code: "Z0", text: "Approval" },
                { code: "Z9", text: "Contracting" },
                { code: "Z3", text: "Final Review" },
                { code: "Z4", text: "Closure" }
            ],
            "MIGP - Dedicated Array": [
                { code: "Z5", text: "Origination" },
                { code: "Z6", text: "Introduction" },
                { code: "Z7", text: "Discovery" },
                { code: "Z8", text: "Structuring" },
                { code: "Z0", text: "Approval" },
                { code: "Z9", text: "Contracting" },
                { code: "Z3", text: "Final Review" },
                { code: "Z4", text: "Closure" }
            ]
        },

        // ─────────────────────────────────────────────────────────────────────
        // _onRouteMatched
        // ─────────────────────────────────────────────────────────────────────
        _onRouteMatched: async function (oEvent) {
            const roles = await this.getUserDetails();
            let oUserScopeJModel = this.getOwnerComponent().getModel("oUserScopeJModel").getData();;
            if (oUserScopeJModel.hasUserAccess || oUserScopeJModel.hasAdminAccess) {
                this._loadOppIdSuggestions();
                try {
                    var oSearchParams = new URLSearchParams(window.location.search);
                    var sUrlOppId = oSearchParams.get("OpportunityID");
                    var sUrlQuoteId = oSearchParams.get("QuoteID");

                    // ── Store BusinessPartnerID from URL ──────────────────────
                    var sBPId = oSearchParams.get("BusinessPartnerID");
                    if (sBPId && sBPId.trim()) {
                        var oSCDetail = this.getView().getModel("oSCDetail");
                        oSCDetail.setProperty("/businessPartnerId", sBPId.trim());
                        console.log("BusinessPartnerID from URL:", sBPId.trim());
                    }

                    // ── Fallback: check hash fragment ─────────────────────────
                    if (!sUrlOppId) {
                        var sHash = window.location.hash || "";
                        var oMatch = sHash.match(/[?&]OpportunityID=([^&]+)/i);
                        if (oMatch) { sUrlOppId = decodeURIComponent(oMatch[1]); }
                    }
                    if (!sUrlQuoteId) {
                        var sHashQ = window.location.hash || "";
                        var oMatchQ = sHashQ.match(/[?&]QuoteID=([^&]+)/i);
                        if (oMatchQ) { sUrlQuoteId = decodeURIComponent(oMatchQ[1]); }
                    }

                    // ── Case 1: OpportunityID only → load Opportunity ─────────
                    if (sUrlOppId && sUrlOppId.trim() && (!sUrlQuoteId || !sUrlQuoteId.trim())) {
                        console.log("Case 1 — Opportunity:", sUrlOppId.trim());
                        var oInput = this.getView().byId("il8");
                        if (oInput) { oInput.setValue(sUrlOppId.trim()); }
                        this.onFetchOpp();
                        return;
                    }

                    // ── Case 2: Both IDs → load/replicate Quote ───────────────
                    if (sUrlOppId && sUrlOppId.trim() && sUrlQuoteId && sUrlQuoteId.trim()) {
                        console.log("Case 2 — Quote (both IDs):", sUrlOppId.trim(), sUrlQuoteId.trim());
                        var oInputC2 = this.getView().byId("il8");
                        if (oInputC2) { oInputC2.setValue(sUrlOppId.trim()); }
                        this._loadOrReplicateQuote(sUrlOppId.trim(), sUrlQuoteId.trim());
                        return;
                    }

                    // ── Case 3: QuoteID only → fetch Quote from DB ────────────
                    if (sUrlQuoteId && sUrlQuoteId.trim() && (!sUrlOppId || !sUrlOppId.trim())) {
                        console.log("Case 3 — Quote (QuoteID only):", sUrlQuoteId.trim());
                        this._loadQuoteByQuoteId(sUrlQuoteId.trim());
                        return;
                    }

                } catch (e) { console.error("Error parsing URL parameters", e); }

                // ── No URL params — fall back to latest CPI record ────────────
                var sCurrentVal = this.getView().byId("il8").getValue();
                if (!sCurrentVal) { this._loadLatestCPIRecord(); }
            }
            else {
                MessageBox.error("No role assigned to your account. Please contact administrator."); return false;
            }
        },


        // ─────────────────────────────────────────────────────────────────────
        // _loadOrReplicateQuote
        // Case 2: Called when C4C passes both OpportunityID and QuoteID
        //   - QuoteID exists in DB → load it directly
        //   - QuoteID missing      → replicate all 5 tables from OppID
        // ─────────────────────────────────────────────────────────────────────
        _loadOrReplicateQuote: async function (sOppId, sQuoteId) {
            var oModel = this.getOwnerComponent().getModel();
            var oSCDetail = this.getView().getModel("oSCDetail");

            // Store context flags — save will write to Quote, not Opp
            oSCDetail.setProperty("/quoteId", sQuoteId);
            oSCDetail.setProperty("/isQuoteContext", true);
            oSCDetail.setProperty("/objectStatus", "QUOTE");

            // Keep OppID in input field (Oppid is same for both Opp and Quote)
            var oInput = this.getView().byId("il8");
            if (oInput) { oInput.setValue(sOppId); }

            try {
                var oOp = oModel.bindContext("/getOrReplicateQuote(...)");
                oOp.setParameter("oppId", sOppId);
                oOp.setParameter("quoteId", sQuoteId);
                await oOp.execute();

                var oResult = oOp.getBoundContext().getObject();
                var oOuter;
                try {
                    oOuter = typeof oResult.value === "string"
                        ? JSON.parse(oResult.value)
                        : oResult.value;
                } catch (e) {
                    MessageBox.error("Unexpected response from server.");
                    return;
                }

                if (!oOuter || oOuter.value === "error") {
                    MessageBox.error("Failed to load/replicate quote: " +
                        (oOuter && oOuter.message ? oOuter.message : "Unknown error"));
                    return;
                }

                var oParsed = typeof oOuter.value === "string"
                    ? JSON.parse(oOuter.value)
                    : oOuter.value;

                if (!oParsed) {
                    MessageBox.error("No data found for Quote: " + sQuoteId);
                    return;
                }

                this._populateAllFields(oParsed);
                MessageBox.success(
                    oParsed._isNew
                        ? "Quote " + sQuoteId + " created from Opportunity " + sOppId + "."
                        : "Quote " + sQuoteId + " loaded successfully."
                );

            } catch (e) {
                MessageBox.error("Failed to load quote: " + e.message);
            }
        },

        // ─────────────────────────────────────────────────────────────────────
        // _loadQuoteByQuoteId
        // Case 3: C4C sends only QuoteID in URL
        // CPI already pushed Quote shell (Oppid=456, objectStatus=QUOTE, quoteId=123)
        // BTP finds shell by quoteId, copies data from parent Opp if first open
        // ─────────────────────────────────────────────────────────────────────
        _loadQuoteByQuoteId: async function (sQuoteId) {
            var oModel = this.getOwnerComponent().getModel();
            var oSCDetail = this.getView().getModel("oSCDetail");

            // Set Quote context flags
            oSCDetail.setProperty("/quoteId", sQuoteId);
            oSCDetail.setProperty("/isQuoteContext", true);
            oSCDetail.setProperty("/objectStatus", "QUOTE");

            try {
                var oOp = oModel.bindContext("/getQuoteByQuoteId(...)");
                oOp.setParameter("quoteId", sQuoteId);
                await oOp.execute();

                var oResult = oOp.getBoundContext().getObject();
                var oOuter;
                try {
                    oOuter = typeof oResult.value === "string"
                        ? JSON.parse(oResult.value)
                        : oResult.value;
                } catch (e) {
                    MessageBox.error("Unexpected response from server.");
                    return;
                }

                if (!oOuter || oOuter.value === "error") {
                    MessageBox.error(
                        "Failed to load Quote: " +
                        (oOuter && oOuter.message ? oOuter.message : "Unknown error")
                    );
                    return;
                }

                var oParsed = typeof oOuter.value === "string"
                    ? JSON.parse(oOuter.value)
                    : oOuter.value;

                if (!oParsed) {
                    MessageBox.error("No data found for Quote: " + sQuoteId);
                    return;
                }

                // Set OppID in input field from DB result (Oppid = 456)
                var oInput = this.getView().byId("il8");
                if (oInput) { oInput.setValue(oParsed.Oppid || ''); }

                this._populateAllFields(oParsed);
                MessageToast.show(
                    oParsed._isNew
                        ? "Quote " + sQuoteId + " created from Opportunity " + oParsed.Oppid + "."
                        : "Quote " + sQuoteId + " loaded successfully."
                );

            } catch (e) {
                MessageBox.error("Failed to load Quote: " + e.message);
            }
        },
        // ─────────────────────────────────────────────────────────────────────
        // _loadOppIdSuggestions
        // ─────────────────────────────────────────────────────────────────────
        _loadOppIdSuggestions: async function () {
            var oModel = this.getOwnerComponent().getModel();
            var oSCDetail = this.getView().getModel("oSCDetail");
            try {
                var oOp = oModel.bindContext("/getAllOppIds(...)");
                await oOp.execute();
                var oResult = oOp.getBoundContext().getObject();
                var aRows;
                try {
                    aRows = typeof oResult.value === "string"
                        ? JSON.parse(oResult.value)
                        : oResult.value;
                } catch (e) { aRows = []; }
                if (!Array.isArray(aRows)) { aRows = []; }
                oSCDetail.setProperty("/OppIdSuggestions", aRows);
            } catch (err) {
                console.warn("getAllOppIds failed:", err.message);
            }
        },

        // ─────────────────────────────────────────────────────────────────────
        // _loadLatestCPIRecord
        // ─────────────────────────────────────────────────────────────────────
        _loadLatestCPIRecord: async function () {
            var oModel = this.getOwnerComponent().getModel();
            try {
                var oOp = oModel.bindContext("/getLatestCPIRecord(...)");
                await oOp.execute();
                var oResult = oOp.getBoundContext().getObject();
                var oOuter;
                try {
                    oOuter = typeof oResult.value === "string"
                        ? JSON.parse(oResult.value)
                        : oResult.value;
                } catch (e) { return; }
                if (!oOuter || oOuter.value === "none") { return; }
                var oParsed = typeof oOuter.value === "string"
                    ? JSON.parse(oOuter.value)
                    : oOuter.value;
                if (!oParsed) { return; }
                this._populateAllFields(oParsed);
                MessageToast.show(
                    "Auto-loaded latest CPI record: " + oParsed.Oppid +
                    (oParsed.modifiedAt
                        ? "  (pushed " + new Date(oParsed.modifiedAt).toLocaleString() + ")"
                        : "")
                );
            } catch (err) {
                console.warn("getLatestCPIRecord failed:", err.message);
            }
        },

        // ─────────────────────────────────────────────────────────────────────
        // _populateAllFields
        // ─────────────────────────────────────────────────────────────────────
        // _populateAllFields
        // ─────────────────────────────────────────────────────────────────────
        _populateAllFields: function (oParsed) {
            var oSCDetail = this.getView().getModel("oSCDetail");
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");

            // ── RESET all fields to editable first ──────────────────────────
            this._unlockOpportunityFields();
            // ────────────────────────────────────────────────────────────────

            oSCDetail.setProperty("/Oppid", oParsed.Oppid || '');
            oSCDetail.setProperty("/quoteId", oParsed.quoteId || '');
            oSCDetail.setProperty("/objectStatus", oParsed.objectStatus || '');
            oSCDetail.setProperty("/isQuoteContext", oParsed.objectStatus === 'QUOTE');
            oSCDetail.setProperty("/oppCategory", oParsed.oppCategory || '');
            oSCDetail.setProperty("/annualGross", oParsed.annualGross || '');
            oSCDetail.setProperty("/sizeofOpp", oParsed.sizeofOpp || '');
            oSCDetail.setProperty("/percentUsage", oParsed.percentUsage || '85');
            oSCDetail.setProperty("/annualSubs", oParsed.annualSubs || '');
            oSCDetail.setProperty("/oppSubs", oParsed.oppSubs || '');
            oSCDetail.setProperty("/contLifetimeval", oParsed.contLifetimeval || '');
            oSCDetail.setProperty("/salesType", oParsed.salesType || '');
            oSCDetail.setProperty("/priceStructure", oParsed.priceStructure || '');
            oSCDetail.setProperty("/nte", oParsed.nte || '');
            oSCDetail.setProperty("/term", oParsed.term || '');
            oSCDetail.setProperty("/meteredCon", oParsed.meteredCon || '');
            oSCDetail.setProperty("/unmeteredCon", oParsed.unmeteredCon || '');
            oSCDetail.setProperty("/enrolled", oParsed.enrolled || '');
            oSCDetail.setProperty("/commencementLetterSent",
                oParsed.commencementLetterSent === true || oParsed.commencementLetterSent === 'true');
            oSCDetail.setProperty("/commencementLetterSigned",
                oParsed.commencementLetterSigned === true || oParsed.commencementLetterSigned === 'true');
            oSCDetail.setProperty("/includeInPipeline",
                oParsed.includeInPipeline === true || oParsed.includeInPipeline === 'true');

            // ── Sales Phase — set from DB value, do NOT reset ────────────────
            var sPhaseFromDB = oParsed.salesPhase || '';
            oSCDetail.setProperty("/salesPhase", sPhaseFromDB);
            oSCDetail.setProperty("/enrollmentReferralCode", oParsed.enrollmentReferralCode || '');
            oSCDetail.setProperty("/enrollmentEmailId", oParsed.enrollmentEmailId || '');
            // ADD after:  oSCDetail.setProperty("/enrollmentEmailId", oParsed.enrollmentEmailId || '');
            oSCDetail.setProperty("/lastFetchedAt", oParsed.lastFetchedAt
                ? new Date(oParsed.lastFetchedAt).toLocaleString("en-US", {
                    month: "short", day: "2-digit", year: "numeric",
                    hour: "2-digit", minute: "2-digit", second: "2-digit"
                })
                : "");

            oSCDetail.setProperty("/lastFetchedAtISO", oParsed.lastFetchedAt || "");

            var oOppInput = this.getView().byId("il8");
            if (oOppInput) { oOppInput.setValue(oParsed.Oppid || ''); }

            // // ── Status ───────────────────────────────────────────────────────
            // var aValidStatuses = [
            //     "Open", "In Progress", "Won", "Lost", "Closed", "Structuring", "In Approval",
            //     "Approved", "Rejected", "Contract Sign-Off"
            // ];
            // var sRawStatus = (oParsed.status || '').trim();
            // var sMatchedStatus = aValidStatuses.find(function (k) {
            //     return k.toLowerCase() === sRawStatus.toLowerCase();
            // }) || '';
            // ── Status ───────────────────────────────────────────────────────────
            var oStatusCodeMap = {
                'Z0': 'Open',
                'Z1': 'In Progress',
                'Z2': 'Won',
                'Z3': 'Lost',
                'Z4': 'Closed',
                'Z5': 'Structuring'
            };
            var aValidStatuses = [
                "Open", "In Progress", "Won", "Lost", "Closed", "Structuring", "In Approval",
                "Approved", "Rejected", "Contract Sign-Off"
            ];
            var sRawStatus = (oParsed.status || '').trim();

            // Resolve Z-code first, then fall through to text match
            var sResolvedStatus = oStatusCodeMap[sRawStatus] || sRawStatus;

            var sMatchedStatus = aValidStatuses.find(function (k) {
                return k.toLowerCase() === sResolvedStatus.toLowerCase();
            }) || '';

            if (sRawStatus && !sMatchedStatus) {
                console.warn("Unknown status value received:", sRawStatus);
                MessageToast.show("Warning: Unknown status '" + sRawStatus + "' received from C4C.");
            }
            oSCDetail.setProperty("/status", sMatchedStatus);

            var oStatusInput = this.getView().byId("idstatus");
            if (oStatusInput) {
                oStatusInput.setValue(sMatchedStatus);
                if (!sMatchedStatus && sRawStatus) {
                    oStatusInput.setValueState("Error");
                    oStatusInput.setValueStateText("Invalid status: '" + sRawStatus + "'");
                } else {
                    oStatusInput.setValueState("None");
                    oStatusInput.setValueStateText("");
                }
            }

            // ── Lock if Approved / Closed / Contract Sign-Off ────────────────
            var aLockedStatuses = ["Approved", "Closed", "Contract Sign-Off", "Won"];
            var bShouldLock = aLockedStatuses.indexOf(sMatchedStatus) !== -1;
            var bIsOpp = !oSCDetail.getProperty("/isQuoteContext");
            if (bShouldLock && bIsOpp) {
                setTimeout(function () {
                    this._lockOpportunityFields(sMatchedStatus);
                }.bind(this), 300);
            }

            // ── ComboBoxes ───────────────────────────────────────────────────
            var oCatCB = this.getView().byId("idModel_CS");
            // if (oCatCB && oParsed.oppCategory) {
            //     oCatCB.setSelectedKey(oParsed.oppCategory);
            //     this.onComboBoxtype();
            // }
            if (oCatCB && oParsed.oppCategory) {
                oCatCB.setSelectedKey(oParsed.oppCategory);
                this.onComboBoxtype();
            }

            // Re-apply Export button visibility now that isQuoteContext is correctly set
            // var oUserScope = this.getOwnerComponent().getModel("oUserScopeJModel");
            // var bAdminNow = oUserScope && oUserScope.oData && oUserScope.oData.hasAdminAccess;
            // var bIsQuoteNow = this.getView().getModel("oSCDetail").getProperty("/isQuoteContext");
            // var oExportBtn = this.getView().byId("idbtnExport");
            // if (oExportBtn) {
            //     oExportBtn.setVisible(!!bAdminNow && !!bIsQuoteNow);
            //     oExportBtn.setEnabled(!!bAdminNow && !!bIsQuoteNow);
            // }
            // ── changed 1 only export buttton for admins not columns ───────────────────────────────────────────────────
            // Re-apply Export button visibility now that isQuoteContext is correctly set
            // var oUserScope = this.getOwnerComponent().getModel("oUserScopeJModel");
            // var bAdminNow = oUserScope && oUserScope.oData && oUserScope.oData.hasAdminAccess;
            // var bIsQuoteNow = this.getView().getModel("oSCDetail").getProperty("/isQuoteContext");
            // var oExportBtn = this.getView().byId("idbtnExport");
            // if (oExportBtn) {
            //     var sCatNow = oSCDetail.getProperty("/oppCategory");
            //     var bShowExport = !!bAdminNow && (
            //         (sCatNow === "MIGP - Small Business" && !bIsQuoteNow) ||
            //         ((sCatNow === "MIGP - LCVP" || sCatNow === "MIGP - Dedicated Array") && bIsQuoteNow)
            //     );
            //     oExportBtn.setVisible(bShowExport);
            //     oExportBtn.setEnabled(bShowExport);
            // }
            // ──change 2 both export buttton and columns visible for admins   ───────────────────────────────────────────────────
            // Re-apply Export button + Export columns visibility (admin + category + context)
            var oUserScope = this.getOwnerComponent().getModel("oUserScopeJModel");
            var bAdminNow = oUserScope && oUserScope.oData && oUserScope.oData.hasAdminAccess;
            var bIsQuoteNow = oSCDetail.getProperty("/isQuoteContext");
            var sCatNow = oSCDetail.getProperty("/oppCategory");

            var bShowExport = !!bAdminNow && (
                (sCatNow === "MIGP - Small Business" && !bIsQuoteNow) ||
                ((sCatNow === "MIGP - LCVP" || sCatNow === "MIGP - Dedicated Array") && bIsQuoteNow)
            );

            // ── Export button ─────────────────────────────────────────────────────────
            var oExportBtn = this.getView().byId("idbtnExport");
            if (oExportBtn) {
                oExportBtn.setVisible(bShowExport);
                oExportBtn.setEnabled(bShowExport);
            }

            // ── SMB table: Export checkbox column + Export Date column ────────────────
            var oSMBExportChkCol = this.getView().byId("_IDGenColumn30b");
            var oSMBExportDateCol = this.getView().byId("_IDGenColumn31");
            if (oSMBExportChkCol) { oSMBExportChkCol.setVisible(bShowExport); }
            if (oSMBExportDateCol) { oSMBExportDateCol.setVisible(bShowExport); }

            // ── LCVP table: Export checkbox column + Export Date column ───────────────
            var oLCVPExportChkCol = this.getView().byId("_IDGenColumn45b");
            var oLCVPExportDateCol = this.getView().byId("_IDGenColumn46");
            if (oLCVPExportChkCol) { oLCVPExportChkCol.setVisible(bShowExport); }
            if (oLCVPExportDateCol) { oLCVPExportDateCol.setVisible(bShowExport); }

            // ── Select-All export checkboxes in column headers ────────────────────────
            var oSelectAllSMB = this.getView().byId("idSelectAllExportSMB");
            var oSelectAllLCVP = this.getView().byId("idSelectAllExportLCVP");
            if (oSelectAllSMB) { oSelectAllSMB.setVisible(bShowExport); }
            if (oSelectAllLCVP) { oSelectAllLCVP.setVisible(bShowExport); }
            // ──End of change 2 both export buttton and columns visible for admins   ───────────────────────────────────────────────────


            var oPriceCB = this.getView().byId("idMode_CS");
            if (oPriceCB && oParsed.priceStructure) { oPriceCB.setSelectedKey(oParsed.priceStructure); }
            var oSalesCB = this.getView().byId("idMode_S");
            if (oSalesCB && oParsed.salesType) { oSalesCB.setSelectedKey(oParsed.salesType); }

            // ── Checkboxes ───────────────────────────────────────────────────
            var oChkSent = this.getView().byId("idcommLetterSent");
            var oChkSigned = this.getView().byId("idcommLetterSigned");
            var oChkPipeline = this.getView().byId("idincludePipeline");
            if (oChkSent) { oChkSent.setSelected(oSCDetail.getProperty("/commencementLetterSent")); }
            if (oChkSigned) { oChkSigned.setSelected(oSCDetail.getProperty("/commencementLetterSigned")); }
            if (oChkPipeline) { oChkPipeline.setSelected(oSCDetail.getProperty("/includeInPipeline")); }

            // ── Tab tables ───────────────────────────────────────────────────
            if (oTabModel) {
                var aIP = oParsed.involvedParties || [];
                var aPC = (oParsed.providerContracts || []).map(function (r) {
                    return Object.assign({}, r, {
                        enrollChk: r.enrollChk === true || r.enrollChk === 'true'
                    });
                });
                var aPR = oParsed.prospects || [];
                var aCD = (oParsed.consumptionDetails || []).map(function (r) {
                    var bSel = r.selected === true || r.selected === 'true';
                    if (r.inactive === 'Yes' || r.inactive === true || !!r.replacementCA) {
                        bSel = false;
                    }
                    return Object.assign({}, r, { selected: bSel });
                });

                console.log('Populating tabs: IP=' + aIP.length + ' PC=' + aPC.length +
                    ' CD=' + aCD.length + ' PR=' + aPR.length);

                oTabModel.setProperty("/InvolvedParties", aIP);
                oTabModel.setProperty("/ProviderContracts", aPC);
                oTabModel.setProperty("/ConsumptionDetails", aCD);
                setTimeout(function () { this._sortCASelectedFirst(); }.bind(this), 100);
                oTabModel.setProperty("/Prospects", aPR);
            }

            // ── Sales Phase ComboBox — load using DB value (FIX: do NOT reset to '') ──
            // var bIsOppCat = !oSCDetail.getProperty("/isQuoteContext");
            // if (bIsOppCat) {
            // var sCatNew = this.getView().byId("idModel_CS").getSelectedKey();
            // ✅ Pass actual salesPhase from DB — never reset to ''
            // this._loadSalesPhaseItems(sCatNew, sPhaseFromDB);
            // }
            // ── Sales Phase ComboBox — always load for both OPP and QUOTE ────
            // _bFieldsPopulated = false so onComboBoxtype doesn't interfere during load
            this._bFieldsPopulated = false;
            var sCatNew = this.getView().byId("idModel_CS").getSelectedKey();
            this._loadSalesPhaseItems(sCatNew, sPhaseFromDB);
            this._bFieldsPopulated = true; // now user interactions can trigger reload

            this._applyOneTimeSaleRestrictions();
            this.onSpecialsExist(true);
        },
        // ─────────────────────────────────────────────────────────────────────
        // _lockOpportunityFields — lock all fields when status = Approved/Closed
        // ─────────────────────────────────────────────────────────────────────
        _lockOpportunityFields: function (sStatus) {

            // ── 1. Header / simple fields (same as before) ────────────────────────
            var aLockIds = [
                "idModel_CS",           // Opportunity Category
                "idMode_CS",            // Price Structure
                "idMode_S",             // Sales Type
                "idnte",                // NTE Price
                "idincludePipeline",    // Include in Pipeline
                "idcommLetterSent",     // Commencement Letter Sent
                "idcommLetterSigned",   // Commencement Letter Signed
                "idsalesPhase",          // Sales Phase
                "idpusage"
            ];
            aLockIds.forEach(function (sId) {
                var oCtrl = this.getView().byId(sId);
                if (oCtrl && oCtrl.setEditable) { oCtrl.setEditable(false); }
                if (oCtrl && oCtrl.setEnabled) { oCtrl.setEnabled(false); }
            }.bind(this));

            // ── 2. Disable Save / Fetch / Calculate buttons ───────────────────────
            var oSaveBtn = this.getView().byId("_IDButton_postgresav");
            if (oSaveBtn) { oSaveBtn.setEnabled(false); }

            var oFetchBtn = this.getView().byId("IDverify");
            if (oFetchBtn) { oFetchBtn.setEnabled(false); }

            var oCalcBtn = this.getView().byId("IDcalculate");
            if (oCalcBtn) { oCalcBtn.setEnabled(false); }

            // ── 3. Consumption Tab — freeze Select-All header checkbox + row checkboxes ──
            // 3a. "Select All" header checkbox
            var oSelectAllCA = this.getView().byId("idSelectAllCA");
            if (oSelectAllCA) { oSelectAllCA.setEnabled(false); }

            // 3b. Row-level checkboxes (template items — iterate table items)
            var oCATable = this.getView().byId("tableISUCA");
            if (oCATable) {
                oCATable.getItems().forEach(function (oItem) {
                    var aCells = oItem.getCells();
                    // Cell 0 is the selected CheckBox
                    if (aCells[0] && aCells[0].setEnabled) {
                        aCells[0].setEnabled(false);
                    }
                });
            }

            // ── 4. Provider Contract Tab — freeze all editable cells in both tables ──
            // Indices of editable cells in SMB table (tablesmall):
            //   0=product, 1=portfolio, 2=fixedPrice, 3=netPremium,
            //   4=portfolioprice, 5=subspercent(Select), 6=startDate,
            //   8=exportChk(CheckBox)
            // ── 4. Provider Contract — SMB table ─────────────────────────────────
            var oSMBTable = this.getView().byId("tablesmall");
            if (oSMBTable) {
                oSMBTable.getItems().forEach(function (oItem) {
                    oItem.getCells().forEach(function (oCell) {
                        if (oCell.setEditable) { oCell.setEditable(false); }
                        if (oCell.setEnabled) { oCell.setEnabled(false); }
                    });
                });
            }
            var oSelectAllSMB = this.getView().byId("idSelectAllExportSMB");
            if (oSelectAllSMB) { oSelectAllSMB.setEnabled(false); }

            // ── Provider Contract — LCVP table ───────────────────────────────────
            var oLCVPTable = this.getView().byId("tablelarge");
            if (oLCVPTable) {
                oLCVPTable.getItems().forEach(function (oItem) {
                    oItem.getCells().forEach(function (oCell) {
                        if (oCell.setEditable) { oCell.setEditable(false); }
                        if (oCell.setEnabled) { oCell.setEnabled(false); }
                    });
                });
            }
            var oSelectAllLCVP = this.getView().byId("idSelectAllExportLCVP");
            if (oSelectAllLCVP) { oSelectAllLCVP.setEnabled(false); }

            // ── 5. Prospect Tab — freeze all row input fields ─────────────────────
            var oProspectTable = this.getView().byId("table11");
            if (oProspectTable) {
                oProspectTable.getItems().forEach(function (oItem) {
                    oItem.getCells().forEach(function (oCell) {
                        if (oCell.setEditable) { oCell.setEditable(false); }
                        if (oCell.setEnabled) { oCell.setEnabled(false); }
                    });
                });
            }
            var oAddRowBtn = this.getView().byId("idbtn_addnewRow_Fuses");
            if (oAddRowBtn) { oAddRowBtn.setEnabled(false); }

            // ── 6. Toast ──────────────────────────────────────────────────────────
            MessageToast.show(
                sStatus === "Won"
                    ? "This Opportunity is Won and locked. Open the Quote to make amendments."
                    : "This Opportunity is locked. Create a Quote to make amendments."
            );
        },

        // ─────────────────────────────────────────────────────────────────────────────
        // _unlockOpportunityFields — restore all fields to editable state
        // Called at START of _populateAllFields to reset before applying lock
        // ─────────────────────────────────────────────────────────────────────────────
        _unlockOpportunityFields: function () {

            // ── 1. Header / simple fields ─────────────────────────────────────────
            var aLockIds = [
                "idModel_CS",
                "idMode_CS",
                "idMode_S",
                "idnte",
                "idincludePipeline",
                "idcommLetterSent",
                "idcommLetterSigned",
                "idsalesPhase",
                "idpusage"
            ];
            aLockIds.forEach(function (sId) {
                var oCtrl = this.getView().byId(sId);
                if (oCtrl && oCtrl.setEditable) { oCtrl.setEditable(true); }
                if (oCtrl && oCtrl.setEnabled) { oCtrl.setEnabled(true); }
            }.bind(this));

            // ── 2. Re-enable Save / Fetch / Calculate buttons ─────────────────────
            var oSaveBtn = this.getView().byId("_IDButton_postgresav");
            if (oSaveBtn) { oSaveBtn.setEnabled(true); }

            var oFetchBtn = this.getView().byId("IDverify");
            if (oFetchBtn) { oFetchBtn.setEnabled(true); }

            var oCalcBtn = this.getView().byId("IDcalculate");
            if (oCalcBtn) { oCalcBtn.setEnabled(true); }

            // ── 3. Consumption Tab — re-enable Select-All header + row checkboxes ──
            var oSelectAllCA = this.getView().byId("idSelectAllCA");
            if (oSelectAllCA) { oSelectAllCA.setEnabled(true); }

            var oCATable = this.getView().byId("tableISUCA");
            if (oCATable) {
                oCATable.getItems().forEach(function (oItem) {
                    var aCells = oItem.getCells();
                    if (aCells[0] && aCells[0].setEnabled) {
                        aCells[0].setEnabled(true);
                    }
                });
            }

            // ── 4. Provider Contract — SMB table ─────────────────────────────────
            var oSMBTable = this.getView().byId("tablesmall");
            if (oSMBTable) {
                oSMBTable.getItems().forEach(function (oItem) {
                    oItem.getCells().forEach(function (oCell) {
                        if (oCell.setEditable) { oCell.setEditable(true); }
                        if (oCell.setEnabled) { oCell.setEnabled(true); }
                    });
                });
            }
            var oSelectAllSMB = this.getView().byId("idSelectAllExportSMB");
            if (oSelectAllSMB) { oSelectAllSMB.setEnabled(true); }

            // ── Provider Contract — LCVP table ───────────────────────────────────
            var oLCVPTable = this.getView().byId("tablelarge");
            if (oLCVPTable) {
                oLCVPTable.getItems().forEach(function (oItem) {
                    oItem.getCells().forEach(function (oCell) {
                        if (oCell.setEditable) { oCell.setEditable(true); }
                        if (oCell.setEnabled) { oCell.setEnabled(true); }
                    });
                });
            }
            var oSelectAllLCVP = this.getView().byId("idSelectAllExportLCVP");
            if (oSelectAllLCVP) { oSelectAllLCVP.setEnabled(true); }

            // ── 5. Prospect Tab — re-enable all row input fields ──────────────────
            var oProspectTable = this.getView().byId("table11");
            if (oProspectTable) {
                oProspectTable.getItems().forEach(function (oItem) {
                    oItem.getCells().forEach(function (oCell) {
                        if (oCell.setEditable) { oCell.setEditable(true); }
                        if (oCell.setEnabled) { oCell.setEnabled(true); }
                    });
                });
            }
            var oAddRowBtn = this.getView().byId("idbtn_addnewRow_Fuses");
            if (oAddRowBtn) { oAddRowBtn.setEnabled(true); }
        },

        // ─────────────────────────────────────────────────────────────────────
        // _applyInactiveFilter — default filter: hide Inactive = 'Yes' rows
        // ─────────────────────────────────────────────────────────────────────
        _applyInactiveFilter: function () {
            var oTable = this.getView().byId("tableISUCA");
            if (!oTable) { return; }
            var oBinding = oTable.getBinding("items");
            if (!oBinding) { return; }
            oBinding.filter([
                new sap.ui.model.Filter("inactive", sap.ui.model.FilterOperator.NE, "Yes")
            ]);
        },

        // ─────────────────────────────────────────────────────────────────────
        // _applyOneTimeSaleRestrictions — Export disabled for One-time Sale (FDD #51)
        // ─────────────────────────────────────────────────────────────────────
        _applyOneTimeSaleRestrictions: function () {
            var oSalesCB = this.getView().byId("idMode_S");
            var oSCDetail = this.getView().getModel("oSCDetail");
            var sSalesType = (oSalesCB && oSalesCB.getSelectedKey())
                ? oSalesCB.getSelectedKey()
                : (oSCDetail ? oSCDetail.getProperty("/salesType") : "");
            var bIsOneTime = (sSalesType === "One-time Sale");
            var oExportBtn = this.getView().byId("idbtnExport");
            if (oExportBtn) { oExportBtn.setEnabled(!bIsOneTime); }
        },

        // ─────────────────────────────────────────────────────────────────────
        // formatThousands — thousands place commas display
        // ─────────────────────────────────────────────────────────────────────
        formatThousands: function (sValue) {
            if (sValue === null || sValue === undefined || sValue === '') return '';
            var fVal = parseFloat(sValue);
            if (isNaN(fVal)) return sValue;
            return fVal.toLocaleString('en-US', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
        },

        // ─────────────────────────────────────────────────────────────────────
        // onProviderContractInputChange — numeric + max 2 decimals validation
        // Used by: portfolio, netPremium
        // ─────────────────────────────────────────────────────────────────────
        onProviderContractInputChange: function (oEvent) {
            var oInput = oEvent.getSource();
            var sValue = oInput.getValue().trim();


            // Allow empty
            if (!sValue) {
                oInput.setValueState("None");
                oInput.setValueStateText("");
                return;
            }

            // Must be a number with max 2 decimal places
            var sNormalized = sValue.replace(/,/g, "");
            var rNumeric = /^\d+(\.\d{1,2})?$/;
            if (!rNumeric.test(sNormalized)) {
                oInput.setValueState("Error");
                oInput.setValueStateText("Enter a valid number with up to 2 decimal places (e.g. 12 or 12.50)");
                //  oInput.setValue("");   // clear invalid value
            } else {
                oInput.setValueState("Success");
                oInput.setValueStateText("");
            }
        },

        // ─────────────────────────────────────────────────────────────────────
        // onFixedOrFlexChange — numeric validation + Fixed vs Flex mutual exclusion
        // Used by: fixedPrice, portfolioprice
        // ─────────────────────────────────────────────────────────────────────
        onFixedOrFlexChange: function (oEvent) {
            var oInput = oEvent.getSource();
            var sValue = oInput.getValue().trim();

            if (sValue) {
                var rNumeric = /^\d{1,3}(,\d{3})*(\.\d{1,2})?$|^\d+(\.\d{1,2})?$/;
                if (!rNumeric.test(sValue)) {
                    oInput.setValueState("Error");
                    oInput.setValueStateText("Enter a valid number with up to 2 decimal places");
                    // oInput.setValue("");
                    return;
                } else {
                    oInput.setValueState("Success");
                    oInput.setValueStateText("");
                }
            } else {
                oInput.setValueState("None");
                oInput.setValueStateText("");
            }
        },

        // ─────────────────────────────────────────────────────────────────────
        // onOppIdSuggest
        // ─────────────────────────────────────────────────────────────────────
        onOppIdSuggest: function (oEvent) {
            var sQuery = (oEvent.getParameter("suggestValue") || "").toLowerCase();
            var oSCDetail = this.getView().getModel("oSCDetail");
            var aAll = oSCDetail.getProperty("/OppIdSuggestions") || [];
            var aFiltered = sQuery
                ? aAll.filter(function (r) {
                    return r.Oppid && r.Oppid.toLowerCase().indexOf(sQuery) !== -1;
                })
                : aAll.slice(0, 20);
            var oInput = this.getView().byId("il8");
            oInput.destroySuggestionItems();
            aFiltered.forEach(function (r) {
                oInput.addSuggestionItem(new sap.ui.core.Item({
                    key: r.Oppid,
                    text: r.Oppid +
                        (r.oppCategory ? "  |  " + r.oppCategory : "") +
                        (r.status ? "  [" + r.status + "]" : "")
                }));
            });
        },

        onOppIdSuggestionSelected: function (oEvent) {
            var oItem = oEvent.getParameter("selectedItem");
            if (!oItem) { return; }
            this.getView().byId("il8").setValue(oItem.getKey());
            this.onFetchOpp();
        },

        // ─────────────────────────────────────────────────────────────────────
        // onFetchOpp — triggered by Enter key (submit) or suggestion select
        // ─────────────────────────────────────────────────────────────────────
        onFetchOpp: async function () {

            var sOppId = this.getView().byId("il8").getValue().trim();
            if (!sOppId) { MessageBox.error("Please enter an Opportunity ID."); return; }
            var oModel = this.getOwnerComponent().getModel();
            var oOperation = oModel.bindContext("/getOpportunityRecord(...)");
            oOperation.setParameter("connection_object", sOppId);
            try {
                await oOperation.execute();
                var oResult = oOperation.getBoundContext().getObject();
                var oOuter;
                try {
                    oOuter = typeof oResult.value === "string"
                        ? JSON.parse(oResult.value) : oResult.value;
                } catch (e) { MessageBox.error("Unexpected response from server."); return; }

                if (!oOuter || oOuter.value === "error" || !oOuter.value) {
                    MessageBox.warning("Could not load opportunity: " +
                        ((oOuter && oOuter.message) ? oOuter.message : "No record found"));
                    return;
                }
                var oParsed = typeof oOuter.value === "string"
                    ? JSON.parse(oOuter.value) : oOuter.value;
                if (!oParsed) { MessageBox.warning("No record found for: " + sOppId); return; }

                this._populateAllFields(oParsed);
                MessageBox.success("Opportunity " + sOppId + " loaded successfully.");
            } catch (err) {
                console.error("getOpportunityRecord failed:", err.message);
                MessageBox.error("Failed to fetch opportunity: " + err.message);
            }

        },

        onSpecialsExist: function (bExists) { console.log("onSpecialsExist:", bExists); },

        // ─────────────────────────────────────────────────────────────────────
        // onComboBoxtype — show/hide LCVP fields + correct PC table
        // ─────────────────────────────────────────────────────────────────────
        onComboBoxtype: function () {
            var sCategory = this.getView().byId("idModel_CS").getSelectedKey();
            var bIsLCVP = (sCategory === "MIGP - LCVP" || sCategory === "MIGP - Dedicated Array");
            var bIsSMB = (sCategory === "MIGP - Small Business" || sCategory === "MIGP Small Business");

            this.getView().byId("tablesmall").setVisible(bIsSMB);
            this.getView().byId("tablelarge").setVisible(bIsLCVP);
            var oSBox = this.getView().byId("idSummaryLineBox");
            if (oSBox) { oSBox.setVisible(false); }

            ["idlblunmtrcon", "idunmtrcon", "idlblmtrcon", "idmtrcon", "idlterm", "idterm",
                // "idlblcommLetterSent", "idcommLetterSent", "idlblcommLetterSigned", "idcommLetterSigned"
            ].forEach(function (sId) {
                var oCtrl = this.getView().byId(sId);
                if (oCtrl) { oCtrl.setVisible(bIsLCVP); }
            }.bind(this));

            // Commencement letter controls: LCVP category AND admin role required
            var oUserScopeJModel = this.getOwnerComponent().getModel("oUserScopeJModel");
            var bIsAdmin = oUserScopeJModel && oUserScopeJModel.oData && oUserScopeJModel.oData.hasAdminAccess;
            console.log("onComboBoxtype - oData:", JSON.stringify(oUserScopeJModel && oUserScopeJModel.oData), "bIsAdmin:", bIsAdmin);
            ["idlblcommLetterSent", "idcommLetterSent", "idlblcommLetterSigned", "idcommLetterSigned"
            ].forEach(function (sId) {
                var oCtrl = this.getView().byId(sId);
                if (oCtrl) { oCtrl.setVisible(bIsLCVP && !!bIsAdmin); }
            }.bind(this));

            // SMB-only fields — Referral Code + Enrollment Email ID
            ["idlblReferralCode", "idReferralCode", "idlblEnrollmentEmail", "idEnrollmentEmail"
            ].forEach(function (sId) {
                var oCtrl = this.getView().byId(sId);
                if (oCtrl) { oCtrl.setVisible(bIsSMB); }
            }.bind(this));

            this._applyOneTimeSaleRestrictions();
            // Change 6 — onComboBoxtype ADD at end:
            // var bIsOppLocal = !this.getView().getModel("oSCDetail")
            //     .getProperty("/isQuoteContext");
            // if (bIsOppLocal) {
            //     var sCat = this.getView().byId("idModel_CS").getSelectedKey();
            //     var sCurrentPhase = this.getView().getModel("oSCDetail").getProperty("/salesPhase") || '';
            //     this._loadSalesPhaseItems(sCat,sCurrentPhase);
            //         // sCat === "MIGP - LCVP" ? "Structuring" : "Approval");
            // }
            // ── When user manually changes category, reload phase items with safe default ──
            // _populateAllFields handles the load-time call with the real DB phase.
            // This block only fires for user-driven category changes.
            if (this._bFieldsPopulated) {
                var sCatSel = this.getView().byId("idModel_CS").getSelectedKey();
                var sPhaseNow = this.getView().getModel("oSCDetail").getProperty("/salesPhase") || '';
                this._loadSalesPhaseItems(sCatSel, sPhaseNow);
            }
        },


        _setLastFetchTime: function () {
            var oModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var now = new Date();
            var sFormatted = now.toLocaleDateString("en-US", {
                month: "short", day: "2-digit", year: "numeric"
            }) + "  " + now.toLocaleTimeString("en-US", {
                hour: "2-digit", minute: "2-digit", second: "2-digit"
            });
            oModel.setProperty("/lastFetchTime", sFormatted);
            oModel.setProperty("/lastFetchVisible", true);
        },
        _persistLastFetchedAt: async function () {
            var oSCDetail = this.getView().getModel("oSCDetail");
            var sOppId = oSCDetail.getProperty("/Oppid");
            var sObjectStatus = oSCDetail.getProperty("/objectStatus") || "OPP";   // ← ADD
            var sQuoteId = oSCDetail.getProperty("/quoteId") || "";                // ← ADD
            if (!sOppId) { return; }

            var oModel = this.getOwnerComponent().getModel();
            try {
                var oOp = oModel.bindContext("/updateLastFetchedAt(...)");
                oOp.setParameter("oppId", sOppId);
                oOp.setParameter("objectStatus", sObjectStatus);   // ← ADD
                oOp.setParameter("quoteId", sQuoteId);             // ← ADD
                await oOp.execute();

                var oResult = oOp.getBoundContext().getObject();
                var oParsed = typeof oResult.value === "string"
                    ? JSON.parse(oResult.value) : oResult.value;

                if (oParsed && oParsed.lastFetchedAt) {
                    var sFormatted = new Date(oParsed.lastFetchedAt).toLocaleString("en-US", {
                        month: "short", day: "2-digit", year: "numeric",
                        hour: "2-digit", minute: "2-digit", second: "2-digit"
                    });
                    oSCDetail.setProperty("/lastFetchedAt", sFormatted);
                    oSCDetail.setProperty("/lastFetchedAtISO", oParsed.lastFetchedAt);
                }
            } catch (e) {
                console.warn("Could not persist lastFetchedAt:", e.message);
            }
        },
        // _setLastValidateTime: function () {
        //     var oModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
        //     var now = new Date();
        //     var sFormatted = now.toLocaleDateString("en-US", {
        //         month: "short", day: "2-digit", year: "numeric"
        //     }) + "  " + now.toLocaleTimeString("en-US", {
        //         hour: "2-digit", minute: "2-digit", second: "2-digit"
        //     });
        //     oModel.setProperty("/lastValidateTime", sFormatted);
        //     oModel.setProperty("/lastValidateVisible", true);
        // },

        // =====================================================================
        // oncalc  --  MIGP Sales Calculation
        // =====================================================================
        //
        // INPUTS (read from UI before any calculation):
        //   grossAnnualUsage  -- SUM(est12month) across ALL CA rows (selected or not)
        //   subPercent        -- header Subscription % (user input, default 85)
        //   nte               -- NTE Price (header field, user input, e.g. 60)
        //   priceStructure    -- "REC Blend" | "Full Price"  (drives REC blend gate)
        //   Per PC row:
        //     subspercent     -- row-level enrollment % (e.g. 50, 85)
        //     portfolioprice  -- Subscription Fee $/MWh   (Flex row)
        //     fixedMWh        -- Fixed MWh amount          (Fixed row / One-time Sale)
        //     recPrice        -- REC Price $/MWh            (required when REC blend)
        //     startDate       -- contract start (YYYY-MM-DD)
        //     endDate         -- contract end   (YYYY-MM-DD)
        //
        // -----------------------------------------------------------------------
        // SHARED STEPS  (run for BOTH SMB and LCVP)
        // -----------------------------------------------------------------------
        //
        // STEP 1  Gross Annual Usage  (FDD #1)
        //         = SUM( CA.est12month )  for ALL CAs in ConsumptionDetails
        //           regardless of whether CA is selected or not
        //         stored in: oSCDetail>/annualGross
        //
        //         Example:  20,000 + 30,000 + 30,000 + 50,000 + 50,000 + 30,000
        //                   = 210,000 MWh  (all 6 CAs)
        //                   scenarios.txt uses selected-CAs-only = 100,000 MWh
        //                   (mock data: selected 3 CAs sum to 100,000)
        //
        // STEP 2  Subscription %  (FDD #2)
        //         = percentUsage (header ComboBox, default 85, valid 1-100)
        //         stored in: oSCDetail>/percentUsage  (user input, not calculated)
        //
        // STEP 3  Size of Opportunity  (FDD #3)
        //         = grossAnnualUsage x subPercent / 100
        //         stored in: oSCDetail>/sizeofOpp
        //
        //         Example:  100,000 x 85 / 100 = 85,000 MWh
        //
        // STEP 4  Metered / Unmetered split  (FDD #14, #15 -- LCVP display only)
        //         meteredTotal   = SUM( est12month ) where CA.metered = "Yes"
        //         unmeteredTotal = SUM( est12month ) where CA.metered != "Yes"
        //         stored in: oSCDetail>/meteredCon, oSCDetail>/unmeteredCon
        //
        // -----------------------------------------------------------------------
        // SMB PATH  (branches off after STEP 4, returns early)
        // -----------------------------------------------------------------------
        //
        // SMB Annual Subscribed Load  (FDD #5)
        //   = SUM( CA.est12month x subPercent/100 )  for SELECTED CAs only
        //   stored in: oSCDetail>/annualSubs
        //
        //   Example:  (20,000 + 50,000 + 30,000) x 85/100
        //             = 100,000 x 0.85 = 85,000 MWh
        //
        // SMB Annual MWh at Close  (FDD #4)
        //   = annualSubs  (snapshot at Calculate; locks after Close Won)
        //   stored in: oSCDetail>/annualMWh
        //
        // SMB Per Provider Contract row:
        //
        //   estUsage  (Flex row -- subspercent entered):
        //     = grossAnnualUsage x row.subspercent / 100
        //     stored in: oOpportunityjmodel>estUsage  (read-only display)
        //
        //   estUsage  (Fixed $ row -- only fixedPrice entered, no subspercent):
        //     = sizeOfOpp  (header-level calc used as MWh proxy)
        //     stored in: oOpportunityjmodel>estUsage  (read-only display)
        //
        //   Term per row  (FDD #10):
        //     termYears = (Date(endDate) - Date(startDate)) / (1000 x 60 x 60 x 24 x 365)
        //
        //   Opportunity Subscribed Load  (FDD #9):
        //     oppSubs = SUM( row.estUsage x row.termYears )
        //     stored in: oSCDetail>/oppSubs
        //
        //     Example:  estUsage=85,000  term=6yr  => oppSubs = 510,000 MWh
        //
        //   Contract Lifetime Value  (FDD #8):
        //     Flex row:   CLV contribution = portfolioprice x estUsage x termYears
        //     Fixed $ row: CLV contribution = fixedPrice    x estUsage x termYears
        //     Total CLV = SUM( priceForCLV x row.estUsage x row.termYears )
        //     stored in: oSCDetail>/contLifetimeval
        //
        //     Example:  Sub Fee=$55, estUsage=85,000, term=6yr
        //               CLV = 55 x 85,000 x 6 = $28,050,000
        //
        //   Total Term  (FDD #10):
        //     = SUM( row.termYears )  across all SMB PC rows
        //     stored in: oSCDetail>/term
        //
        // -----------------------------------------------------------------------
        // LCVP PATH  (continues from STEP 4)
        // -----------------------------------------------------------------------
        //
        // STEP 5  Validation
        //         - Portfolio required on every row
        //         - Each row must have EITHER fixedMWh OR portfolioprice, not both
        //         - startDate and endDate required on every row
        //         - recPrice required when priceStructure="REC Blend" AND SubFee > NTE
        //
        // STEP 6  Term per row  (FDD #10)
        //         row._termYears = (Date(endDate) - Date(startDate))
        //                          / (1000 x 60 x 60 x 24 x 365)
        //         totalTermYears = SUM( row._termYears )
        //         stored in: oSCDetail>/term = totalTermYears.toFixed(2)
        //
        //         Example -- Scenario 4:
        //           Row 1: 2027-01-01 to 2030-12-31 = 4.00 years
        //           Row 2: 2031-01-01 to 2033-12-31 = 3.00 years
        //           totalTermYears = 7.00 years
        //
        // STEP 7  Per-row calculations  (LCVP Flex + REC Blend logic)
        //
        //   7a. Est. Usage per row  (Flex rows only):
        //       estUsage = grossAnnualUsage x row.subspercent / 100
        //       stored in: oOpportunityjmodel>estUsage  (read-only)
        //
        //       NOTE: grossAnnualUsage is used directly -- NOT sizeOfOpp.
        //             Each row has its own sub%, applied independently to gross.
        //
        //       Example -- Scenario 4:
        //         Row 1: 100,000 x 50/100 = 50,000 MWh
        //         Row 2: 100,000 x 85/100 = 85,000 MWh
        //
        //   7b. REC Blend decision (evaluated PER ROW independently):
        //       Triggers when BOTH are true:
        //         (i)  priceStructure = "REC Blend"   (header dropdown)
        //         (ii) row.portfolioprice > NTE        (Sub Fee exceeds NTE on this row)
        //
        //       If Sub Fee <= NTE on a row => that row is 100% MIGP
        //       even when other rows in same quote trigger REC blend.
        //
        //   7c. REC Blend row formulas:
        //       REC MWh  = estUsage x (NTE - SubFee) / (RECPrice - SubFee)
        //       MIGP MWh = estUsage - REC MWh
        //       MIGP %   = MIGP MWh / estUsage x 100
        //       REC %    = REC MWh  / estUsage x 100
        //       _calcMWh = MIGP MWh  (used for header aggregations -- FDD #8 note)
        //
        //       Example -- Scenario 4 Row 1 (Sub=$75, NTE=$60, REC=$5, estUsage=50,000):
        //         REC MWh  = 50,000 x (60-75) / (5-75)
        //                  = 50,000 x (-15/-70)
        //                  = 50,000 x 0.21429 = 10,714.29 MWh
        //         MIGP MWh = 50,000 - 10,714.29  = 39,285.71 MWh
        //         MIGP %   = 39,285.71 / 50,000  x 100 = 78.57%
        //         REC %    = 10,714.29 / 50,000  x 100 = 21.43%
        //         _calcMWh = 39,285.71
        //
        //       Example -- Scenario 4 Row 2 (Sub=$85, NTE=$60, REC=$6, estUsage=85,000):
        //         REC MWh  = 85,000 x (60-85) / (6-85)
        //                  = 85,000 x (-25/-79)
        //                  = 85,000 x 0.31646 = 26,898.73 MWh
        //         MIGP MWh = 85,000 - 26,898.73  = 58,101.27 MWh
        //         MIGP %   = 58,101.27 / 85,000  x 100 = 68.35%
        //         REC %    = 26,898.73 / 85,000  x 100 = 31.65%
        //         _calcMWh = 58,101.27
        //
        //   7d. No REC Blend row (SubFee <= NTE or Full Price structure):
        //       MIGP MWh = estUsage,  REC MWh = 0
        //       MIGP %   = 100,       REC %   = 0
        //       _calcMWh = estUsage
        //
        //   7e. Fixed MWh row (One-time Sale / LCVP Fixed):
        //       No estUsage (sub% not applicable for Fixed MWh)
        //       No REC blend  (Fixed MWh rows always 100% MIGP)
        //       _calcMWh = fixedMWh  (user-entered Fixed MWh amount)
        //
        // STEP 8  Header aggregations  (weighted by term)
        //
        //   weight per row = row._termYears / totalTermYears
        //
        //   Annual Subscribed Load  (FDD #5):
        //     = SUM( weight x row._calcMWh )
        //     = SUM( (row._termYears / totalTermYears) x row._calcMWh )
        //     stored in: oSCDetail>/annualSubs
        //
        //     Example -- Scenario 4:
        //       = (4/7 x 39,285.71) + (3/7 x 58,101.27)
        //       = 22,448.98 + 24,900.54 = 47,349.52 MWh
        //
        //   Opportunity Subscribed Load  (FDD #9):
        //     = SUM( row._termYears x row._calcMWh )
        //     stored in: oSCDetail>/oppSubs
        //
        //     Example -- Scenario 4:
        //       = (4 x 39,285.71) + (3 x 58,101.27)
        //       = 157,142.84 + 174,303.81 = 331,446.65 MWh
        //
        //   Contract Lifetime Value  (FDD #8):
        //     = SUM( row.portfolioprice x row._calcMWh x row._termYears )
        //     uses _calcMWh (= MIGP MWh when REC blend) per FDD 5.3 note #8
        //     stored in: oSCDetail>/contLifetimeval
        //
        //     Example -- Scenario 4:
        //       = (75 x 39,285.71 x 4) + (85 x 58,101.27 x 3)
        //       = 11,785,713.00 + 14,815,823.85 = $26,601,536.85
        //
        //   Annual MWh at Close  (FDD #4):
        //     = annualSubscribedLoad at the time of Calculate
        //     locks after status = "Close Won" (no recalculation after that)
        //     stored in: oSCDetail>/annualMWh
        //
        // STEP 9  Summary Bar  (idSummaryLineBox HBox in View1.view.xml)
        //         Shown below tablelarge only when 2+ LCVP rows exist.
        //         Hidden for single-row scenarios (1, 2, 8).
        //         All 4 values are weighted averages using row term as weight.
        //
        //   weight per row = row._termYears / totalTermYears
        //
        //   Weighted Avg Sub Fee  (summarySubFee):
        //     = SUM( weight x row.portfolioprice )
        //     stored in: oSCDetail>/summarySubFee
        //
        //     Example -- Scenario 4:
        //       = (4/7 x 75) + (3/7 x 85) = 42.86 + 36.43 = $79.29
        //
        //   Weighted Avg Sub %  (summarySubPct):
        //     = SUM( weight x row.subspercent )
        //     stored in: oSCDetail>/summarySubPct
        //
        //     Example -- Scenario 4:
        //       = (4/7 x 50) + (3/7 x 85) = 28.57 + 36.43 = 65.00%
        //
        //   Weighted Avg MIGP %  (summaryMigpPct):
        //     = SUM( weight x row.migpPercent )
        //     uses CALCULATED migpPercent from STEP 7, not user input
        //     stored in: oSCDetail>/summaryMigpPct
        //
        //     Example -- Scenario 4:
        //       = (4/7 x 78.57) + (3/7 x 68.35) = 44.90 + 29.29 = 74.19%
        //
        //   Weighted Avg REC %  (summaryRecPct):
        //     = SUM( weight x row.recpercent )
        //     uses CALCULATED recpercent from STEP 7, not user input
        //     stored in: oSCDetail>/summaryRecPct
        //
        //     Example -- Scenario 4:
        //       = (4/7 x 21.43) + (3/7 x 31.65) = 12.24 + 13.56 = 25.81%
        //
        // -----------------------------------------------------------------------
        // SCENARIO QUICK REFERENCE  (from scenarios.txt)
        // -----------------------------------------------------------------------
        //
        //  Sc  Escalation  REC Blend  Portfolio  Key behaviour
        //  --  ----------  ---------  ---------  -------------------------------------------
        //  1   No          No         No         Sub $55 < NTE $60 => 100% MIGP, 1 row
        //  2   No          Yes        No         Sub $75 > NTE $60 => REC blend, 1 row
        //  3   Yes         Yes        No         2 rows same portfolio P1, both REC blend
        //  4   Yes         Yes        Yes        P1(4yr,$75,50%) + P2(3yr,$85,85%), REC blend
        //  5   Yes         No         Yes        P1 Sub<NTE + P2 Sub>NTE, Full Price => 100%
        //  6   No          Yes        Yes        P1+P2 overlapping dates OK (diff portfolios)
        //  7   No          No         Yes        P1+P2 overlapping dates, Full Price => 100%
        //  8   Fixed MWh   N/A        No         MIGP_R17_FXD_MWh, One-time Sale, no Export
        //
        // =====================================================================
        oncalc: function () {

            // ── Validate portfolioprice (Flex $/MWh) in LCVP table ───────────
            var bHasError = false;
            var rNumeric = /^\d{1,3}(,\d{3})*(\.\d{1,2})?$|^\d+(\.\d{1,2})?$/;

            var oLCVPTable = this.getView().byId("tablelarge");
            if (oLCVPTable && oLCVPTable.getVisible()) {
                oLCVPTable.getItems().forEach(function (oItem) {
                    var oCell = oItem.getCells()[2]; // portfolioprice is cell index 2
                    if (!oCell || !oCell.getValue) { return; }
                    var sRaw = oCell.getValue().trim();
                    var sClean = sRaw.replace(/,/g, "");
                    if (!sRaw) { return; } // allow empty
                    if (isNaN(sClean) || !rNumeric.test(sRaw)) {
                        oCell.setValueState("Error");
                        oCell.setValueStateText("Enter a valid number with up to 2 decimal places");
                        bHasError = true;
                    } else {
                        oCell.setValueState("None");
                    }
                });
            }

            // ── Validate numeric fields in SMB table ─────────────────────────────
            var oSMBTable = this.getView().byId("tablesmall");
            if (oSMBTable && oSMBTable.getVisible()) {
                oSMBTable.getItems().forEach(function (oItem) {
                    var aCells = oItem.getCells();
                    // cell 2=fixedPrice, 3=netPremium, 4=portfolioprice
                    [2, 3, 4].forEach(function (iIdx) {
                        var oCell = aCells[iIdx];
                        if (!oCell || !oCell.getValue) { return; }
                        var sRaw = oCell.getValue().trim();
                        var sClean = sRaw.replace(/,/g, "");
                        if (!sRaw) { return; } // allow empty
                        if (isNaN(sClean) || !rNumeric.test(sRaw)) {
                            oCell.setValueState("Error");
                            oCell.setValueStateText("Enter a valid number with up to 2 decimal places");
                            bHasError = true;
                        } else {
                            oCell.setValueState("None");
                        }
                    });
                });
            }

            if (bHasError) {
                MessageToast.show("Please fix Subscription Fee ($/MWh) — only numbers up to 2 decimals allowed.");
                return; // ← blocks calculation
            }

            // ─────────────────────────────────────────────────────────────────


            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var oSCDetail = this.getView().getModel("oSCDetail");
            if (!oTabModel || !oSCDetail) {
                MessageBox.error("Models not initialized. Please fetch opportunity first.");
                return;
            }

            // ── BUG 17907: Block calculate when Opportunity is fully closed ──
            var sCalcStatus = oSCDetail.getProperty("/status") || "";
            var sCalcPhase = oSCDetail.getProperty("/salesPhase") || "";
            if (sCalcPhase === "Z4" && (sCalcStatus === "Closed" || sCalcStatus === "Won" || sCalcStatus === "Lost")) {
                MessageBox.warning(
                    "This Opportunity is in phase Z4 (Closure) with status '" + sCalcStatus +
                    "'. Calculate is not available for closed opportunities."
                );
                return;
            }
            // ────────────────────────────────────────────────────────────────

            var sCategory = this.getView().byId("idModel_CS").getSelectedKey();
            var bIsLCVP = (sCategory === "MIGP - LCVP");

            // ─────────────────────────────────────────────────────────────────────────
            // HELPER: exact calendar-month count between two Date objects.
            //   calendarMonths(new Date("2027-01-01"), new Date("2032-12-31")) → 72
            //   calendarMonths(new Date("2027-01-01"), new Date("2027-01-31")) → 1  (Sc8)
            // Logic: full years × 12 + remaining month difference,
            //        then +1 if the end day ≥ start day (inclusive end date convention).
            // [FIX-TERM] replaces all (dEnd-dStart)/(MS_PER_DAY*365.25) usages.
            // ─────────────────────────────────────────────────────────────────────────
            function calendarMonths(dStart, dEnd) {
                if (!dStart || !dEnd || isNaN(dStart) || isNaN(dEnd) || dEnd <= dStart) return 0;
                var y1 = dStart.getFullYear(), m1 = dStart.getMonth(), d1 = dStart.getDate();
                var y2 = dEnd.getFullYear(), m2 = dEnd.getMonth(), d2 = dEnd.getDate();
                var months = (y2 - y1) * 12 + (m2 - m1);
                // +1 only when end day is STRICTLY greater than start day
                // (inclusive end convention: Jan 1 → Dec 31 = 12 months)
                // When days are equal it's an exact month/year boundary — no +1 needed
                // e.g. Jan 1 2027 → Jan 1 2030 = exactly 36 months (3 yrs), not 37
                if (d2 > d1) months += 1;
                return months > 0 ? months : 0;
            }

            // ── STEP 1: Gross Annual Usage ────────────────────────────────────────────
            // = SUM(CA.est12month) for ALL CAs  +  SUM(Prospect.projectedCon) for ALL Prospects
            var aCARows = oTabModel.getProperty("/ConsumptionDetails") || [];
            var aProspectRows = oTabModel.getProperty("/Prospects") || [];
            var grossAnnualUsage = 0;
            var selectedCAsTotal = 0;

            // CA contribution
            aCARows.forEach(function (oCA) {
                var usage = parseFloat(oCA.est12month) || 0;
                grossAnnualUsage += usage;
                if (oCA.selected === true || oCA.selected === 'true') {
                    selectedCAsTotal += usage;
                }
            });

            // Prospect contribution (projectedCon field)
            aProspectRows.forEach(function (oPR) {
                var projUsage = parseFloat(oPR.projectedCon) || 0;
                grossAnnualUsage += projUsage;
            });

            oSCDetail.setProperty("/annualGross", grossAnnualUsage.toFixed(2));

            // ── STEP 2: Subscription % ────────────────────────────────────────────────
            var subPercent = parseFloat(oSCDetail.getProperty("/percentUsage")) || 85;
            if (subPercent < 1 || subPercent > 100) {
                MessageBox.error("Estimated Subscription Percentage must be between 1 and 100.");
                return;
            }

            // ── STEP 3: Size of Opportunity ──────────────────────────────────────────
            var sizeOfOpp = grossAnnualUsage * (subPercent / 100);
            oSCDetail.setProperty("/sizeofOpp", sizeOfOpp.toFixed(2));


            // ── STEP 4: Metered / Unmetered totals ───────────────────────────────────
            var meteredTotal = 0, unmeteredTotal = 0;
            aCARows.forEach(function (oCA) {
                var usage = parseFloat(oCA.est12month) || 0;
                if ((oCA.metered || '').toLowerCase() === "yes") {
                    meteredTotal += usage;
                } else {
                    unmeteredTotal += usage;
                }
            });
            oSCDetail.setProperty("/meteredCon", meteredTotal.toFixed(2));
            oSCDetail.setProperty("/unmeteredCon", unmeteredTotal.toFixed(2));

            // ═════════════════════════════════════════════════════════════════════════
            // SMB PATH
            // ═════════════════════════════════════════════════════════════════════════
            if (!bIsLCVP) {
                var aSMBValidRows = oTabModel.getProperty("/ProviderContracts") || [];

                if (aSMBValidRows.length === 0) {
                    MessageBox.error("No Provider Contract rows found. Please add product details first.");
                    return;
                }
                if (aSMBValidRows.length > 1) {
                    MessageBox.error(
                        "SMB allows only 1 product line item in the Provider Contract table. " +
                        "Escalation is not applicable for Small Business (FDD Business Rules)."
                    );
                    return;
                }

                for (var si = 0; si < aSMBValidRows.length; si++) {
                    var oSMBRow = aSMBValidRows[si];
                    var smbHasFixedPrice = parseFloat(oSMBRow.fixedPrice) > 0;
                    var smbHasSubFee = parseFloat(oSMBRow.portfolioprice) > 0;
                    var smbHasSubPct = parseFloat(oSMBRow.subspercent) > 0;

                    if (smbHasSubPct) {
                        var nSmbSubPct = parseFloat(oSMBRow.subspercent);
                        if (nSmbSubPct < 5 || nSmbSubPct > 100 || nSmbSubPct % 5 !== 0) {
                            MessageBox.error("Row " + (si + 1) +
                                ": Subscription % must be between 5 and 100 in multiples of 5 (e.g. 5, 10, 15...).");
                            return;
                        }
                    }
                    // if (!smbHasFixedPrice && !smbHasSubFee && !smbHasSubPct) {
                    //     MessageBox.error("Row " + (si + 1) +
                    //         ": Either Fixed Price or Subscription Fee / Percentage is required.");
                    //     return;
                    // }
                    // Fixed Dollar scenario requires both fixedPrice AND netPremium
                    var smbNetPremium = parseFloat(oSMBRow.netPremium) || 0;
                    var bFixedDollarRow = (smbHasFixedPrice && smbNetPremium > 0);

                    if (smbHasFixedPrice && smbNetPremium <= 0) {
                        MessageBox.error("Row " + (si + 1) +
                            ": Fixed Price enrollment requires a Net Premium $/MWh value.");
                        return;
                    }
                    if (bFixedDollarRow && !smbHasSubFee) {
                        MessageBox.error("Row " + (si + 1) +
                            ": Fixed Price enrollment requires a Subscription Fee $/MWh for CLV calculation.");
                        return;
                    }
                    if (!smbHasFixedPrice && !smbHasSubFee && !smbHasSubPct) {
                        MessageBox.error("Row " + (si + 1) +
                            ": Either Fixed Price (with Net Premium + Subscription Fee) or Subscription Fee / Percentage is required.");
                        return;
                    }
                    if (!oSMBRow.startDate) {
                        MessageBox.error("Row " + (si + 1) + ": Start Date is required.");
                        return;
                    }
                }

                // Annual Subscribed Load = selected CAs usage × Sub%
                // Annual Subscribed Load — computed per-row below (Fixed $ uses netPremium formula;
                // Flex uses CA-usage × Sub%). We derive the header value after the row loop.
                var aSMBRows = oTabModel.getProperty("/ProviderContracts") || [];
                var smb_annualSubs = 0;  // accumulated in row loop
                var smb_oppSubs = 0;
                var smb_clv = 0;

                // [FIX-TERM] SMB term: track earliest start and latest end for header display
                var smbEarliestStart = null;
                var smbLatestEnd = null;

                if (aSMBRows.length > 0) {
                    aSMBRows.forEach(function (oRow) {
                        var subPct = parseFloat(oRow.subspercent) || 0;
                        var fixedMWhV = parseFloat(oRow.fixedMWh) || 0;
                        var fixedPrc = parseFloat(oRow.fixedPrice) || 0;
                        var subFee = parseFloat(oRow.portfolioprice) || 0;

                        var dStart = new Date(oRow.startDate);
                        var dEnd = new Date(oRow.endDate);

                        // [FIX-TERM] exact calendar months
                        var rowMonths = calendarMonths(dStart, dEnd);
                        var rowTermYears = rowMonths / 12;
                        oRow._termMonths = rowMonths;

                        // Track header span dates
                        if (!isNaN(dStart) && (!smbEarliestStart || dStart < smbEarliestStart)) smbEarliestStart = dStart;
                        if (!isNaN(dEnd) && (!smbLatestEnd || dEnd > smbLatestEnd)) smbLatestEnd = dEnd;

                        // ── SMB: detect Fixed Dollar scenario ─────────────────────────
                        // Fixed Dollar = fixedPrice > 0 AND netPremium > 0
                        // (fixedMWh is the LCVP Fixed-MWh concept; fixedPrice is the SMB $ amount)
                        var netPremium = parseFloat(oRow.netPremium) || 0;
                        var bFixedDollar = (fixedPrc > 0 && netPremium > 0);

                        // estUsage per row
                        var rowEstUsage = 0;
                        // Count selected CAs once 
                        var nSelectedCAs = aCARows.filter(function (ca) {
                            return ca.selected === true || ca.selected === 'true';
                        }).length;
                        if (nSelectedCAs === 0) nSelectedCAs = 1; // safety fallback avoid zeroing of calculation

                        if (bFixedDollar) {
                            // [SMB-FIXED-$] Annual Subscribed Load = (1 / netPremium) * fixedPrice * 12
                            // estUsage here represents the Annual Subscribed Load for this row
                            // rowEstUsage = (1 / netPremium) * fixedPrc * 12;
                            // [SMB-FIXED-$] fixedPrice is per-CA $/month; multiply by CA count first
                            var totalMonthlyDollars = fixedPrc * nSelectedCAs;
                            rowEstUsage = (1 / netPremium) * totalMonthlyDollars * 12;
                        } else if (fixedMWhV > 0) {
                            rowEstUsage = fixedMWhV;
                        } else if (subPct > 0) {
                            rowEstUsage = selectedCAsTotal * (subPct / 100);
                        } else if (fixedPrc > 0) {
                            // fixedPrice with no netPremium — treat as flat fixed price (original logic)
                            rowEstUsage = selectedCAsTotal * (subPercent / 100);
                        }
                        oRow.estUsage = rowEstUsage > 0 ? rowEstUsage.toFixed(2) : "";
                        smb_annualSubs += rowEstUsage; // annual load contribution from this row


                        if (bFixedDollar) {
                            // [SMB-FIXED-$]
                            // Annual Subscribed Load = rowEstUsage (already computed above)
                            // Opp Subscribed Load   = annualSubsLoad × term
                            // CLV                   = oppSubs × subFee
                            var rowOppSubs = rowEstUsage * rowTermYears;
                            smb_oppSubs += rowOppSubs;
                            smb_clv += rowOppSubs * subFee;
                        } else if (subFee > 0) {
                            smb_oppSubs += rowEstUsage * rowTermYears;
                            smb_clv += subFee * rowEstUsage * rowTermYears;
                        } else if (fixedPrc > 0) {
                            smb_oppSubs += rowEstUsage * rowTermYears;
                            smb_clv += fixedPrc * rowTermYears;
                        }
                    });

                    oTabModel.setProperty("/ProviderContracts", aSMBRows);
                }

                // [FIX-TERM] SMB header term from span (not sum)
                var smbHeaderMonths = calendarMonths(smbEarliestStart, smbLatestEnd);
                var smbDisplayYears = Math.floor(smbHeaderMonths / 12);
                var smbDisplayMonths = smbHeaderMonths % 12;
                var smbTermDisplay = smbDisplayYears + " yr" + (smbDisplayYears !== 1 ? "s" : "") +
                    (smbDisplayMonths > 0 ? " " + smbDisplayMonths + " month" + (smbDisplayMonths !== 1 ? "s" : "") : "");
                oSCDetail.setProperty("/term", smbTermDisplay);

                oSCDetail.setProperty("/annualSubs", smb_annualSubs.toFixed(2));
                oSCDetail.setProperty("/oppSubs", smb_oppSubs.toFixed(2));
                oSCDetail.setProperty("/contLifetimeval", smb_clv.toFixed(2));

                MessageBox.success(
                    "Calculations complete:\n" +
                    "Gross Annual Usage:        " + grossAnnualUsage.toFixed(2) + " MWh\n" +
                    "Size of Opportunity:       " + sizeOfOpp.toFixed(2) + " MWh\n" +
                    "Total Term:               " + smbTermDisplay + "\n" +
                    "Annual Subscribed Load:    " + smb_annualSubs.toFixed(2) + " MWh\n" +
                    "Opp Subscribed Load:       " + smb_oppSubs.toFixed(2) + " MWh\n" +
                    "Contract Lifetime Value:   $" + smb_clv.toFixed(2)
                );
                return;
            }

            // ═════════════════════════════════════════════════════════════════════════
            // LCVP PATH
            // ═════════════════════════════════════════════════════════════════════════
            var aRows = oTabModel.getProperty("/ProviderContracts") || [];

            if (aRows.length === 0) {
                MessageBox.error("No Provider Contract rows found. Please add product details first.");
                return;
            }

            var nte = parseFloat(oSCDetail.getProperty("/nte")) || 0;
            var sPriceStr = this.getView().byId("idMode_CS").getSelectedKey();
            if (!sPriceStr) {
                MessageBox.error("Please select a Price Structure before calculating.");
                return;
            }
            var bRECStructure = (sPriceStr === "REC Blend");

            // ── STEP 5: Validate rows ─────────────────────────────────────────────────
            var oPortfolioMap = {};
            for (var pi = 0; pi < aRows.length; pi++) {
                var oRow = aRows[pi];
                var sPort = oRow.portfolio;
                var dStart = new Date(oRow.startDate);
                var dEnd = new Date(oRow.endDate);

                if (!sPort) { continue; }
                if (!oPortfolioMap[sPort]) { oPortfolioMap[sPort] = []; }

                for (var pj = 0; pj < oPortfolioMap[sPort].length; pj++) {
                    var oExisting = oPortfolioMap[sPort][pj];
                    if (dStart < oExisting.end && dEnd > oExisting.start) {
                        MessageBox.error(
                            "Row " + (pi + 1) + ": Date overlap detected for portfolio '" + sPort +
                            "'. Same portfolio cannot have overlapping dates."
                        );
                        return;
                    }
                }
                oPortfolioMap[sPort].push({ start: dStart, end: dEnd });
            }

            for (var i = 0; i < aRows.length; i++) {
                var oRow = aRows[i];
                if (!oRow.portfolio) {
                    MessageBox.error("Row " + (i + 1) + ": Portfolio is required."); return;
                }
                var hasFixed = parseFloat(oRow.fixedMWh) > 0;
                var hasFlex = parseFloat(oRow.portfolioprice) > 0;
                var sSubPct = oRow.subspercent !== undefined ? String(oRow.subspercent).trim() : "";
                var hasSubPct = sSubPct !== "" && parseFloat(sSubPct) > 0;

                if (hasFixed && hasSubPct) {
                    MessageBox.error("Row " + (i + 1) +
                        ": Fixed MWh row cannot have Subscription % — remove Sub % for Fixed MWh product.");
                    return;
                }
                if (!hasFixed && !hasFlex) {
                    MessageBox.error("Row " + (i + 1) + ": Subscription Fee (Flex) or Fixed MWh is required.");
                    return;
                }
                if (hasSubPct) {
                    var nSubPct = parseFloat(sSubPct);
                    if (nSubPct < 5 || nSubPct > 100 || nSubPct % 5 !== 0) {
                        MessageBox.error("Row " + (i + 1) +
                            ": Subscription % must be between 5 and 100 in multiples of 5 (e.g. 5, 10, 15...).");
                        return;
                    }
                }
                if (!oRow.startDate) {
                    MessageBox.error("Row " + (i + 1) + ": Start Date is required."); return;
                }
                if (!oRow.endDate) {
                    MessageBox.error("Row " + (i + 1) + ": End Date is required for LCVP."); return;
                }
                var subFeeCheck = parseFloat(oRow.portfolioprice) || 0;
                if (bRECStructure && subFeeCheck > nte && !oRow.recPrice) {
                    MessageBox.error("Row " + (i + 1) +
                        ": REC Price is required when Subscription Fee exceeds NTE Price.");
                    return;
                }
            }

            // ── STEP 6: Term per row + header term ───────────────────────────────────
            // [FIX-TERM] Use calendarMonths() — eliminates 365.25 floating-point drift.
            // [FIX-WEIGHT] Header term = span(earliestStart → latestEnd), NOT sum of rows.
            //   This is the KEY fix for Sc6 & Sc7: two overlapping rows each spanning 72 months
            //   still produce a header of 72 months (6 yrs), not 144 months (12 yrs).
            var overallEarliestStart = null;
            var overallLatestEnd = null;

            aRows.forEach(function (oRow) {
                var dStart = new Date(oRow.startDate);
                var dEnd = new Date(oRow.endDate);

                // [FIX-TERM] exact calendar months per row
                oRow._termMonths = calendarMonths(dStart, dEnd);
                oRow._termYears = oRow._termMonths / 12;

                if (!isNaN(dStart) && (!overallEarliestStart || dStart < overallEarliestStart)) {
                    overallEarliestStart = dStart;
                }
                if (!isNaN(dEnd) && (!overallLatestEnd || dEnd > overallLatestEnd)) {
                    overallLatestEnd = dEnd;
                }
            });

            // [FIX-TERM] Header term: exact calendar months across the full contract span
            var termTotalMonths = calendarMonths(overallEarliestStart, overallLatestEnd);
            var totalTermYears = termTotalMonths / 12;

            // Term display
            var termDisplayYears = Math.floor(termTotalMonths / 12);
            var termDisplayMonths = termTotalMonths % 12;
            var termDisplay = termDisplayYears + " yr" + (termDisplayYears !== 1 ? "s" : "") +
                (termDisplayMonths > 0 ? " " + termDisplayMonths + " month" + (termDisplayMonths !== 1 ? "s" : "") : "");

            // [FIX-SC8] One-time sale: 0 yrs 1 month display
            if (termDisplayYears === 0 && termDisplayMonths === 1) {
                termDisplay = "0 yrs 1 month";
            }

            oSCDetail.setProperty("/term", termDisplay);
            oSCDetail.setProperty("/termMonths", termTotalMonths);

            // ── STEP 7: Per-row MWh and REC Blend ────────────────────────────────────
            var errorOccurred = false;
            aRows.forEach(function (oRow, idx) {
                var subFee = parseFloat(oRow.portfolioprice) || 0;
                var subPct = parseFloat(oRow.subspercent) || 0;
                var fixedMWh = parseFloat(oRow.fixedMWh) || 0;
                var recPrice = parseFloat(oRow.recPrice) || 0;

                var estUsage = selectedCAsTotal * (subPct / 100);
                oRow.estUsage = estUsage.toFixed(2);

                if (fixedMWh > 0) {
                    // Fixed MWh product — no REC blend, no subscription %
                    oRow.migpMWh = fixedMWh.toFixed(2);
                    oRow.recMWh = "0.00";
                    oRow.migpPercent = "100.00";
                    oRow.recpercent = "0.00";
                    oRow._calcMWh = fixedMWh;

                } else if (subFee > nte && bRECStructure && recPrice !== subFee) {
                    // REC Blend path
                    var divisor = recPrice - subFee;
                    if (divisor === 0) {
                        MessageBox.error("Row " + (idx + 1) +
                            ": REC Price and Subscription Fee cannot be equal (division by zero).");
                        errorOccurred = true;
                        oRow._calcMWh = 0;
                        return;
                    }
                    var recMWh = estUsage * (nte - subFee) / divisor;
                    var migpMWh = estUsage - recMWh;

                    oRow.recMWh = recMWh.toFixed(2);
                    oRow.migpMWh = migpMWh.toFixed(2);
                    oRow.migpPercent = estUsage > 0 ? ((migpMWh / estUsage) * 100).toFixed(2) : "0.00";
                    oRow.recpercent = estUsage > 0 ? ((recMWh / estUsage) * 100).toFixed(2) : "0.00";
                    oRow._calcMWh = migpMWh;  // header aggregations use MIGP MWh

                } else {
                    // Standard (no REC blend)
                    oRow.migpMWh = estUsage.toFixed(2);
                    oRow.recMWh = "0.00";
                    oRow.migpPercent = "100.00";
                    oRow.recpercent = "0.00";
                    oRow._calcMWh = estUsage;
                }
            });

            if (errorOccurred) { return; }

            // ── STEP 8: Header aggregations ──────────────────────────────────────────
            //
            // [FIX-WEIGHT] Weight = row._termMonths / termTotalMonths  (header span).
            //   For non-overlapping rows this is equivalent to the old logic.
            //   For overlapping rows (Sc6, Sc7) the denominator is the true contract
            //   duration, so weights stay ≤ 1 and Opp Subscribed Load is not doubled.
            //
            // Annual Subscribed Load = Σ (weight_i × MIGP_MWh_i)
            //   "What is the average annual load across the full contract life?"
            //
            // Opp Subscribed Load = Σ (rowTermYears_i × MIGP_MWh_i)
            //   "Total MWh across the entire contract."
            //
            // CLV = Σ (subFee_i × MIGP_MWh_i × rowTermYears_i)
            //   [FIX-CLV-SC4] Per-row multiplication avoids weighted-average distortion.
            //
            // [FIX-SC8] One-time Fixed-MWh sale (termTotalMonths ≤ 1):
            //   oppSubscribedLoad = annualSubscribedLoad (treat as single period).
            //   CLV = subFee × fixedMWh (no term multiplier).
            //
            var annualSubscribedLoad = 0;
            var oppSubscribedLoad = 0;
            var contractLifetimeValue = 0;

            var bOneTimeSale = (termTotalMonths <= 1);  // [FIX-SC8]

            if (totalTermYears > 0 || bOneTimeSale) {

                // ── [FIX-SC8] One-time Fixed-MWh sale ────────────────────────────────
                if (bOneTimeSale) {
                    aRows.forEach(function (r) {
                        if (parseFloat(r.fixedMWh) > 0) {
                            annualSubscribedLoad = r._calcMWh;
                            oppSubscribedLoad = r._calcMWh;
                            contractLifetimeValue = (parseFloat(r.portfolioprice) || 0) * r._calcMWh;
                        }
                    });

                } else {
                    // ── [BUG-A FIX] Compute sumRowTermMonths first ────────────────────
                    // For non-overlapping scenarios: sumRowTermMonths == termTotalMonths → no change.
                    // For overlapping portfolios: sumRowTermMonths > termTotalMonths → weights < 1 → correct.
                    var sumRowTermMonths = 0;
                    aRows.forEach(function (r) {
                        sumRowTermMonths += r._termMonths;
                    });

                    // // ── Pass 1: annualSubs and wAvgSubFee ─────────────────────────────
                    // var wAvgSubFee = 0;
                    // aRows.forEach(function (r) {
                    //     // [BUG-A FIX] use sumRowTermMonths as denominator (not termTotalMonths)
                    //     var weight = sumRowTermMonths > 0 ? (r._termMonths / sumRowTermMonths) : 0;

                    //     annualSubscribedLoad += weight * r._calcMWh;
                    //     wAvgSubFee += weight * (parseFloat(r.portfolioprice) || 0);
                    // });

                    // // ── [BUG-B FIX] oppSubs = annualSubs × totalTermYears ────────────
                    // oppSubscribedLoad = annualSubscribedLoad * totalTermYears;

                    // // ── [BUG-C FIX] CLV = wAvgSubFee × oppSubs ───────────────────────
                    // contractLifetimeValue = wAvgSubFee * oppSubscribedLoad;
                    aRows.forEach(function (r) {
                        var weight = sumRowTermMonths > 0 ? (r._termMonths / sumRowTermMonths) : 0;
                        var rowTermYears = r._termMonths / 12;
                        var rowSubFee = parseFloat(r.portfolioprice) || 0;

                        annualSubscribedLoad += weight * r._calcMWh;

                        // [FIX-CLV] Per-row: oppSubs and CLV accumulate independently
                        var rowOppSubs = r._calcMWh * rowTermYears;
                        oppSubscribedLoad += rowOppSubs;
                        contractLifetimeValue += rowSubFee * rowOppSubs;
                    });
                }
            }

            oSCDetail.setProperty("/annualSubs", annualSubscribedLoad.toFixed(2));
            oSCDetail.setProperty("/oppSubs", oppSubscribedLoad.toFixed(2));
            oSCDetail.setProperty("/contLifetimeval", contractLifetimeValue.toFixed(2));

            // ── STEP 9: Summary bar ──────────────────────────────────────────────────
            var aDisplayRows = aRows.map(function (r) {
                var clone = Object.assign({}, r);
                delete clone._termYears;
                delete clone._calcMWh;
                return clone;
            });

            var oSummaryBox = this.getView().byId("idSummaryLineBox");

            if (aDisplayRows.length > 1 && totalTermYears > 0) {

                // [FIX-WEIGHT] Use sumRowTermMonths as denominator — same fix as Step 8.
                // This keeps the summary-line weights consistent with the header aggregations.
                // For non-overlapping rows sumRowTermMonths == termTotalMonths → no change.
                // For overlapping portfolios (Sc6/Sc7) the weights are correctly normalised.
                var sumRowTermMonths_s9 = 0;
                aRows.forEach(function (r) { sumRowTermMonths_s9 += r._termMonths; });

                var wAvgSubFee = 0;
                var wAvgSubPct = 0;
                var wAvgMIGPPct = 0;
                var wAvgRECPct = 0;
                var wAvgRECPrice = 0;          // [NEW] weighted-average REC Price
                var hasAnyRECPrice = false;    // only show if at least one row has a REC Price

                aRows.forEach(function (r) {
                    var w = sumRowTermMonths_s9 > 0 ? (r._termMonths / sumRowTermMonths_s9) : 0;

                    if (parseFloat(r.portfolioprice) > 0) {
                        wAvgSubFee += w * (parseFloat(r.portfolioprice) || 0);
                    }
                    wAvgSubPct += w * (parseFloat(r.subspercent) || 0);
                    wAvgMIGPPct += w * (parseFloat(r.migpPercent) || 0);
                    wAvgRECPct += w * (parseFloat(r.recpercent) || 0);

                    // [NEW] Include REC Price in weighted average only for rows that have one.
                    // Rows without a REC Price (no-blend rows) contribute 0, which is correct
                    // because their recpercent is also 0 — the weighted avg naturally reflects
                    // only the periods where REC blending is active.
                    var recPriceVal = parseFloat(r.recPrice) || 0;
                    if (recPriceVal > 0) {
                        hasAnyRECPrice = true;
                    }
                    wAvgRECPrice += w * recPriceVal;
                });

                oSCDetail.setProperty("/summarySubFee", wAvgSubFee.toFixed(2));
                oSCDetail.setProperty("/summarySubPct", wAvgSubPct.toFixed(2));
                oSCDetail.setProperty("/summaryMigpPct", wAvgMIGPPct.toFixed(2));
                oSCDetail.setProperty("/summaryRecPct", wAvgRECPct.toFixed(2));
                // [NEW] Set weighted-average REC Price; blank string if no row has a REC Price
                oSCDetail.setProperty("/summaryRecPrice", hasAnyRECPrice ? wAvgRECPrice.toFixed(2) : "");

                if (oSummaryBox) { oSummaryBox.setVisible(true); }

            } else {
                oSCDetail.setProperty("/summarySubFee", "");
                oSCDetail.setProperty("/summarySubPct", "");
                oSCDetail.setProperty("/summaryMigpPct", "");
                oSCDetail.setProperty("/summaryRecPct", "");
                oSCDetail.setProperty("/summaryRecPrice", "");   // [NEW]
                if (oSummaryBox) { oSummaryBox.setVisible(false); }
            }

            oTabModel.setProperty("/ProviderContracts", aDisplayRows);

            aRows.forEach(function (r) {
                delete r._termYears;
                delete r._calcMWh;
            });

            MessageBox.success(
                "Calculations complete:\n\n" +
                "Gross Annual Usage:       " + grossAnnualUsage.toFixed(2) + " MWh\n" +
                "Size of Opportunity:      " + sizeOfOpp.toFixed(2) + " MWh\n" +
                "Total Term:               " + termDisplay + "\n" +
                "Annual Subscribed Load:   " + annualSubscribedLoad.toFixed(2) + " MWh\n" +
                "Opp Subscribed Load:      " + oppSubscribedLoad.toFixed(2) + " MWh\n" +
                "Contract Lifetime Value:  $" + contractLifetimeValue.toFixed(2)
            );
        },

        // =====================================================================
        // onfetchCA
        // Fetch / Validate button on the CA Consumption Details tab.
        //
        // MASTER SWITCH: bUseMock = true  → mock data (CPI not ready)
        //                bUseMock = false → live CPI call (CPI ready)
        //
        // Mock data: 6 CAs for BP 100000010:
        //   900001  — Selected=X  | 20,000 MWh | Metered | clean
        //   900002  — NOT selected | 30,000 MWh | Inactive=Yes | ReplacedBy=9000010
        //   900003  — NOT selected | 30,000 MWh | Metered | Arrear=Yes | NSF=Yes
        //   900004  — Selected=X  | 50,000 MWh | clean
        //   900005  — NOT selected | 50,000 MWh | NSF=Yes
        //   9000010 — Selected=X  | 30,000 MWh | clean (replacement CA for 900002)
        // =====================================================================
        // ─────────────────────────────────────────────────────────────────────
        // onfetchCA
        // Fetch / Validate button on the CA Consumption Details tab.
        //
        // MASTER SWITCH: bUseMock = true  → mock data (CPI not ready)
        //                bUseMock = false → live CPI call (CPI ready)
        //
        // Mock data — simulates what CPI returns for BP 100000010:
        //
        //   FIRST FETCH (aExisting is empty — table was blank):
        //     CPI returns ALL CAs including inactive ones per spec.
        //     900001  — active  | metered
        //     900002  — inactive=Yes | replacedBy=9000010  ← hidden by default filter
        //     900003  — active  | arrear + NSF flags
        //     900004  — active  | clean
        //     900005  — active  | NSF flag
        //     9000010 — active  | replacement for 900002
        //
        //   NEXT FETCH (aExisting has rows — table was already populated):
        //     CPI returns all CURRENTLY active CAs.
        //     900001 is intentionally ABSENT → simulates it became inactive/dropped.
        //     900011 is NEW → simulates a new active CA added for this BP.
        //     _mergeFetchedCAs will:
        //       • 900001 was in existing → not in fresh list → added back as inactive=Yes
        //       • 900011 not in existing → brand-new, unselected
        //       • all others: selection preserved if still active
        // ─────────────────────────────────────────────────────────────────────
        onfetchCA: async function () {
            // ── BUG 17907: Block fetch when Opportunity is fully closed ──────
            var oSCDetailGuard = this.getView().getModel("oSCDetail");
            if (oSCDetailGuard) {
                var sGuardStatus = oSCDetailGuard.getProperty("/status") || "";
                var sGuardPhase = oSCDetailGuard.getProperty("/salesPhase") || "";
                var bIsClosedPhase = (sGuardPhase === "Z4");
                var bIsClosedStatus = (sGuardStatus === "Closed" || sGuardStatus === "Won" || sGuardStatus === "Lost");
                if (bIsClosedPhase && bIsClosedStatus) {
                    MessageBox.warning(
                        "This Opportunity is in phase Z4 (Closure) with status '" + sGuardStatus +
                        "' and is fully locked. Fetch/Validate CA is not available."
                    );
                    return;
                }
            }
            // ────────────────────────────────────────────────────────────────
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            // ── ADD THIS ─────────────────────────────────────────────────────
            console.log("=== onfetchCA DEBUG ===");
            console.log("oTabModel exists:", !!oTabModel);
            if (oTabModel) {
                var oData = oTabModel.getData();
                console.log("InvolvedParties:", JSON.stringify(oData.InvolvedParties));
                console.log("All model keys:", Object.keys(oData));
            }
            console.log("Full model:", JSON.stringify(oTabModel ? oTabModel.getData() : null));
            // ─
            if (!oTabModel) {
                MessageBox.error("Model not initialized. Please fetch the Opportunity first.");
                return;
            }

            var oView = this.getView();
            var oSCDetail = this.getView().getModel("oSCDetail");
            var oModel = this.getOwnerComponent().getModel();
            var aExisting = oTabModel.getProperty("/ConsumptionDetails") || [];
            var bIsFirstFetch = (aExisting.length === 0);

            // ── MASTER SWITCH ────────────────────────────────────────────────
            // Set to true  → always use mock data (CPI not ready)
            // Set to false → use live CPI call    (CPI ready)
            var bUseMock = false; // ← switch to false only when CPI has real CA data for your test BP
            // ────────────────────────────────────────────────────────────────

            oView.setBusy(true);

            // ── MOCK DATA BLOCK ──────────────────────────────────────────────
            if (bUseMock) {
                console.warn("onfetchCA: using mock data — CPI not ready.");
                console.log("onfetchCA: bIsFirstFetch =", bIsFirstFetch,
                    "| aExisting.length =", aExisting.length);

                // ── First Fetch mock ─────────────────────────────────────────
                // Spec: "populate CA table with ALL CAs including Inactive CAs
                //        with Inactive flag = Yes"
                // CPI returns everything — active + inactive — on first fetch.
                // The inactive filter hides inactive rows from the user by default.
                var aMockFirstFetch = [
                    {
                        contacc: "900001", buspartner: "100000010",
                        est12month: "20000", est12monthMetered: "20000", est12monthUnMetered: "0",
                        metered: "Yes", nsfFlag: "", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        // inactive CA — Inactive=Yes + ReplacedBy=9000010
                        // hidden by default filter; 9000010 appears as its replacement below
                        contacc: "900002", buspartner: "100000010",
                        est12month: "30000", est12monthMetered: "0", est12monthUnMetered: "30000",
                        metered: "", nsfFlag: "", arrear60: "", inactive: "Yes", replacementCA: "9000010",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        contacc: "900003", buspartner: "100000010",
                        est12month: "30000", est12monthMetered: "30000", est12monthUnMetered: "0",
                        metered: "Yes", nsfFlag: "Yes", arrear60: "Yes", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        contacc: "900004", buspartner: "100000010",
                        est12month: "50000", est12monthMetered: "0", est12monthUnMetered: "50000",
                        metered: "", nsfFlag: "", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        contacc: "900005", buspartner: "100000010",
                        est12month: "50000", est12monthMetered: "0", est12monthUnMetered: "50000",
                        metered: "", nsfFlag: "Yes", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        // replacement CA for 900002 — active, clean
                        contacc: "9000010", buspartner: "100000010",
                        est12month: "30000", est12monthMetered: "0", est12monthUnMetered: "30000",
                        metered: "", nsfFlag: "", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    }
                ];

                // ── Next Fetch mock ──────────────────────────────────────────
                // Spec: "refresh CA table with CAs received from CR&B"
                // CPI returns all CURRENTLY active CAs for the BPs.
                // 900001 is absent → dropped from CPI (became inactive).
                //   _mergeFetchedCAs will add it back as inactive=Yes regardless
                //   of whether it was selected — so user can see it dropped off.
                // 900011 is new → brand-new CA for this BP, added unselected.
                var aMockNextFetch = [
                    {
                        contacc: "900002", buspartner: "100000010",
                        est12month: "30000", est12monthMetered: "0", est12monthUnMetered: "30000",
                        metered: "", nsfFlag: "", arrear60: "", inactive: "Yes", replacementCA: "9000010",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        contacc: "900003", buspartner: "100000010",
                        est12month: "30000", est12monthMetered: "30000", est12monthUnMetered: "0",
                        metered: "Yes", nsfFlag: "Yes", arrear60: "Yes", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        contacc: "900004", buspartner: "100000010",
                        est12month: "50000", est12monthMetered: "0", est12monthUnMetered: "50000",
                        metered: "", nsfFlag: "", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        contacc: "900005", buspartner: "100000010",
                        est12month: "50000", est12monthMetered: "0", est12monthUnMetered: "50000",
                        metered: "", nsfFlag: "Yes", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        contacc: "9000010", buspartner: "100000010",
                        est12month: "30000", est12monthMetered: "0", est12monthUnMetered: "30000",
                        metered: "", nsfFlag: "", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2026-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    },
                    {
                        // brand-new CA — not in previous fetch, added unselected
                        contacc: "900011", buspartner: "100000010",
                        est12month: "15000", est12monthMetered: "15000", est12monthUnMetered: "0",
                        metered: "Yes", nsfFlag: "", arrear60: "", inactive: "", replacementCA: "",
                        validity: "2027-12-31", providercontractId: "", cycle20: "", enrolled: "", RECBlender_flag: ""
                    }
                ];

                // Pick which mock array to use based on whether table was blank
                var aMockCAs = bIsFirstFetch ? aMockFirstFetch : aMockNextFetch;

                var aResult = this._mergeFetchedCAs(aExisting, aMockCAs);
                oTabModel.setProperty("/ConsumptionDetails", aResult);

                // On first fetch: apply inactive filter so user only sees active CAs
                // On next fetch:  keep whatever filter state the user already set
                if (bIsFirstFetch) {
                    setTimeout(function () { this._applyInactiveFilter(); }.bind(this), 0);
                }

                this._sortCASelectedFirst();
                this.getView().byId("tablesaveCA").setVisible(false);
                this.getView().byId("tableISUCA").setVisible(true);
                oView.setBusy(false);

                // Persist timestamp — only writes lastFetchedAt, not CA data
                this._persistLastFetchedAt();

                MessageToast.show(
                    bIsFirstFetch
                        ? "CA data loaded. Inactive CAs hidden by default — use filter to show."
                        : "CA data refreshed. Selections preserved. Review changes before saving."
                );
                return;
            }
            // ── END MOCK DATA BLOCK ──────────────────────────────────────────

            // ── LIVE CPI CALL — runs when bUseMock = false ───────────────────
            try {
                var sOppId = oSCDetail.getProperty("/Oppid") || "";
                var sObjStatus = oSCDetail.getProperty("/objectStatus") || "OPP";
                var sQuoteId = oSCDetail.getProperty("/quoteId") || "";

                // Collect all BPs from InvolvedParties — deduplicated
                var aAllIPs = oTabModel.getProperty("/InvolvedParties") || [];
                console.log("[fetchCA] InvolvedParties count:", aAllIPs.length);
                console.log("InvolvedParties:", JSON.stringify(aAllIPs)); // ← add this
                console.log("Full model data:", JSON.stringify(oTabModel.getData()));
                var aSeenBPs = [];
                var aBPsToSend = [];
                aAllIPs.forEach(function (ip) {
                    console.log("[fetchCA] IP entry:", JSON.stringify(ip)); // see exact field names
                    if (ip.buspartner && aSeenBPs.indexOf(ip.buspartner) === -1) {
                        aSeenBPs.push(ip.buspartner);
                        aBPsToSend.push({ businessPartnerId: ip.buspartner });
                    }
                });
                console.log("[fetchCA] BPs to send:", JSON.stringify(aBPsToSend));

                if (aBPsToSend.length === 0) {
                    oView.setBusy(false);
                    MessageBox.error(
                        "No Business Partners found in Involved Parties. " +
                        "Please add at least one BP before fetching Contract Accounts."
                    );
                    return;
                }

                // First Fetch → selectedCAs = [] (BPs only sent to CPI)
                // Next Fetch  → selectedCAs = previously selected CA ids
                //               (CPI checks these for inactivity + returns all active CAs)
                var aSelectedCAIds = bIsFirstFetch ? [] : (aExisting || [])
                    .filter(function (ca) { return ca.selected === true || ca.selected === "true"; })
                    .map(function (ca) { return ca.contacc; });

                console.log("onfetchCA: bIsFirstFetch =", bIsFirstFetch,
                    "| BPs =", aBPsToSend.length,
                    "| selectedCAs =", aSelectedCAIds.length);

                var oOp = oModel.bindContext("/fetchValidateCAs(...)");
                oOp.setParameter("businessPartners", JSON.stringify(aBPsToSend));
                oOp.setParameter("selectedCAs", JSON.stringify(aSelectedCAIds));
                oOp.setParameter("oppId", sOppId);
                oOp.setParameter("objectStatus", sObjStatus);
                oOp.setParameter("quoteId", sQuoteId);
                await oOp.execute();

                var oRaw = oOp.getBoundContext().getObject();
                var oFetchResult;
                try {
                    oFetchResult = typeof oRaw.value === "string"
                        ? JSON.parse(oRaw.value) : oRaw.value;
                } catch (e) {
                    oView.setBusy(false);
                    // MessageBox.error("Unexpected response from CPI.");
                    // Parse the error message returned from service.js req.error()
                    var sMsg = (e.message || e.toString() || "");

                    if (sMsg.indexOf("DEST_NOT_FOUND") !== -1) {
                        MessageBox.error(
                            "Destination 'Salescloud_CRB_CPI' could not be resolved.\n\n" +
                            "Check that the destination exists in BTP Cockpit and the app is bound to the Destination Service.",
                            { title: "Destination Not Found" }
                        );
                    } else if (sMsg.indexOf("CPI_AUTH_FAILED") !== -1) {
                        MessageBox.error(
                            "CPI rejected the request — authentication failed.\n\n" +
                            "Check the username and password configured in the BTP Cockpit destination.",
                            { title: "CPI Authentication Failed" }
                        );
                    } else if (sMsg.indexOf("CPI_NOT_FOUND") !== -1) {
                        MessageBox.error(
                            "CPI iFlow URL returned 404 — not found.\n\n" +
                            "Check the URL configured in destination 'Salescloud_CRB_CPI' in BTP Cockpit.",
                            { title: "CPI URL Not Found" }
                        );
                    } else if (sMsg.indexOf("CPI_UNREACHABLE") !== -1) {
                        MessageBox.warning(
                            "No response received from CPI.\n\n" +
                            "The destination URL may be incorrect or CPI is temporarily unavailable. Check BTP Cockpit destination URL.",
                            { title: "CPI Unreachable" }
                        );
                    } else if (sMsg.indexOf("CPI_ERROR") !== -1) {
                        MessageBox.error(
                            "CPI returned an unexpected error.\n\n" + sMsg,
                            { title: "CPI Error" }
                        );
                    } else {
                        // Fallback for any other error
                        MessageBox.error(
                            "Fetch/Validate failed: " + sMsg,
                            { title: "Fetch Failed" }
                        );
                    }
                    return;
                }

                var aFreshCAs = oFetchResult.contractAccounts || [];
                if (!Array.isArray(aFreshCAs)) {
                    oView.setBusy(false);
                    MessageBox.error("Invalid response from CPI — expected array of CAs.");

                    return;
                }
                if (aFreshCAs.length === 0) {
                    oView.setBusy(false);
                    var sKeys = Object.keys(oFetchResult).join(", "); // ← define it here from oFetchResult
                    MessageBox.warning(
                        // "No eligible electric Contract Accounts found " +
                        // "for selected Business Partner(s)."
                        "No Contract Accounts were returned by CPI for Business Partner: " +
                        aBPsToSend.map(function (bp) { return bp.businessPartnerId; }).join(", ") +
                        ".\n\nPossible reasons:\n" +
                        "• BP has no eligible electric Contract Accounts in CR&B\n" +
                        "• BP exists but has only gas accounts\n" +
                        "• CR&B environment has no data for this BP",
                        { title: "No Contract Accounts Found" }
                    );
                    return;
                }

                // Map CPI response fields → CA table fields
                // selected is always set to false here — _mergeFetchedCAs restores
                // the correct selection state from aExisting
                var aMapped = aFreshCAs.map(function (r) {
                    return {
                        contacc: r.contractAccountId || "",
                        buspartner: r.businessPartnerId || "",
                        est12month: r.est12month || "",
                        est12monthMetered: r.est12monthMetered || "",
                        est12monthUnMetered: r.est12monthUnMetered || "",
                        metered: r.metered || "",
                        arrear60: r.arrear60 || "",
                        nsfFlag: r.nsfFlag || "",
                        replacementCA: r.replacementCA || "",
                        inactive: r.inactive || "",
                        enrolled: r.enrolled || "",
                        cycle20: r.cycle20 || "",
                        validity: r.validity || "",
                        providercontractId: r.providercontractId || "",
                        RECBlender_flag: r.RECBlender_flag || "",
                        selected: false
                    };
                });

                var aResult = this._mergeFetchedCAs(aExisting, aMapped);
                oTabModel.setProperty("/ConsumptionDetails", aResult);

                // On first fetch: apply inactive filter so user only sees active CAs
                // On next fetch:  keep whatever filter state the user already set
                if (bIsFirstFetch) {
                    setTimeout(function () { this._applyInactiveFilter(); }.bind(this), 0);
                }

                this._sortCASelectedFirst();
                this.getView().byId("tablesaveCA").setVisible(false);
                this.getView().byId("tableISUCA").setVisible(true);
                oView.setBusy(false);

                // Persist lastFetchedAt timestamp from server response if available
                if (oFetchResult.lastFetchedAt) {
                    oSCDetail.setProperty("/lastFetchedAt",
                        new Date(oFetchResult.lastFetchedAt).toLocaleString("en-US", {
                            month: "short", day: "2-digit", year: "numeric",
                            hour: "2-digit", minute: "2-digit", second: "2-digit"
                        })
                    );
                }

                MessageToast.show(
                    bIsFirstFetch
                        ? "Contract Accounts loaded. Inactive CAs hidden by default — use filter to show."
                        : "Contract Accounts refreshed. Selections preserved. Review changes before saving."
                );

            } catch (e) {
                oView.setBusy(false);
                MessageBox.error("Fetch/Validate failed: " + e.message);
            }
        },

        // ─────────────────────────────────────────────────────────────────────
        // _mergeFetchedCAs
        //
        // Merges CPI fresh data with what the user already has in the table.
        //
        // First Fetch (aExisting empty):
        //   Returns aFreshCAs as-is. All CAs unselected.
        //   Caller applies inactive filter (hides inactive rows by default).
        //
        // Next Fetch (aExisting has rows):
        //   Step 1 — process every CA CPI returned:
        //     • CA exists in table + still active   → preserve selected flag
        //     • CA exists in table + now inactive   → set selected=false (auto-uncheck)
        //     • CA exists in table + has replacementCA → set selected=false (auto-uncheck)
        //     • CA is brand-new (new site)          → add as new row, selected=false
        //
        //   Step 2 — process CAs that CPI did NOT return this time:
        //     • Any CA in existing table that CPI dropped entirely
        //       → add it back with inactive=Yes, selected=false
        //       → this covers BOTH selected and unselected CAs so the user
        //         can see everything that dropped off, regardless of selection
        // ─────────────────────────────────────────────────────────────────────
        _mergeFetchedCAs: function (aExisting, aFreshCAs) {

            // ── First Fetch — nothing to merge ───────────────────────────────
            if (aExisting.length === 0) {
                // Return CPI data as-is, all unselected.
                // Caller applies _applyInactiveFilter to hide inactive rows.
                return aFreshCAs.map(function (ca) {
                    return Object.assign({}, ca, { selected: false });
                });
            }

            // ── Next Fetch — build lookup maps ───────────────────────────────
            var oExistingMap = {};
            aExisting.forEach(function (ca) { oExistingMap[ca.contacc] = ca; });

            var oFreshMap = {};
            aFreshCAs.forEach(function (ca) { oFreshMap[ca.contacc] = ca; });

            // ── Step 1: Process each CA CPI returned ─────────────────────────
            var aResult = aFreshCAs.map(function (oNew) {
                var oOld = oExistingMap[oNew.contacc];

                if (oOld) {
                    // CA was already in the table — decide if selection survives
                    var bWasSelected = (oOld.selected === true || oOld.selected === "true");
                    // Spec: "if CA becomes inactive or is replaced → auto-unselect"
                    var bNowInactive = (oNew.inactive === "Yes" || !!oNew.replacementCA);
                    oNew.selected = bWasSelected && !bNowInactive;

                    // Preserve RECBlender_flag from previous fetch if CPI did not return it
                    if (!oNew.RECBlender_flag) {
                        oNew.RECBlender_flag = oOld.RECBlender_flag || "";
                    }
                } else {
                    // Brand-new CA (new site) — always starts unselected
                    oNew.selected = false;
                }

                return oNew;
            });

            // ── Step 2: CAs that CPI dropped entirely ────────────────────────
            // Spec: "Column Inactive will be set to Yes" for ALL CAs that
            // become inactive — not just previously selected ones.
            // Add every dropped CA back as inactive=Yes so the user can see
            // what changed, then decide whether to save.
            aExisting.forEach(function (oOld) {
                if (!oFreshMap[oOld.contacc]) {
                    // CA is no longer returned by CPI → mark inactive, unselect
                    aResult.push(Object.assign({}, oOld, {
                        inactive: "Yes",
                        selected: false
                    }));
                }
            });

            return aResult;
        },

        // ─────────────────────────────────────────────────────────────────────
        // _sortCASelectedFirst — sorts selected CAs to top of the table
        // ─────────────────────────────────────────────────────────────────────
        _sortCASelectedFirst: function () {
            var oTable = this.getView().byId("tableISUCA");
            if (!oTable) { return; }
            var oBinding = oTable.getBinding("items");
            if (!oBinding) { return; }
            oBinding.sort(new sap.ui.model.Sorter("selected", false,
                false,
                function (a, b) {
                    var valA = (a === true || a === "true") ? 1 : 0;
                    var valB = (b === true || b === "true") ? 1 : 0;
                    return valB - valA;
                }
            ));
        },

        // ─────────────────────────────────────────────────────────────────────
        // onSelectAllCA — Select All / Deselect All checkbox in column header
        // ─────────────────────────────────────────────────────────────────────
        onSelectAllCA: function (oEvent) {
            var bSelected = oEvent.getParameter("selected");
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var aData = oTabModel.getProperty("/ConsumptionDetails") || [];

            var aUpdated = aData.map(function (oCA) {
                var bInactiveOrReplaced = (oCA.inactive === "Yes" || oCA.inactive === true || !!oCA.replacementCA);
                return Object.assign({}, oCA, {
                    selected: bInactiveOrReplaced ? false : bSelected
                });
            });

            oTabModel.setProperty("/ConsumptionDetails", aUpdated);
        },

        // ── onSelectAllExportSMB — Select/Deselect all Export checkboxes in SMB table ──
        onSelectAllExportSMB: function (oEvent) {
            var bSelected = oEvent.getParameter("selected");
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var aPC = oTabModel.getProperty("/ProviderContracts") || [];
            aPC.forEach(function (r) { r.exportChk = bSelected; });
            oTabModel.setProperty("/ProviderContracts", aPC);
        },

        // ── onSelectAllExportLCVP — Select/Deselect all Export checkboxes in LCVP table ──
        onSelectAllExportLCVP: function (oEvent) {
            var bSelected = oEvent.getParameter("selected");
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var aPC = oTabModel.getProperty("/ProviderContracts") || [];
            aPC.forEach(function (r) { r.exportChk = bSelected; });
            oTabModel.setProperty("/ProviderContracts", aPC);
        },

        // ─────────────────────────────────────────────────────────────────────
        // onCARowSelect — individual row checkbox — sync Select All header state
        // ─────────────────────────────────────────────────────────────────────
        onCARowSelect: function () {
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var aData = oTabModel.getProperty("/ConsumptionDetails") || [];

            var aSelectable = aData.filter(function (oCA) {
                return !(oCA.inactive === "Yes" || oCA.inactive === true || !!oCA.replacementCA);
            });
            var bAllSelected = aSelectable.length > 0 && aSelectable.every(function (oCA) {
                return oCA.selected === true || oCA.selected === "true";
            });

            var oSelectAll = this.getView().byId("idSelectAllCA");
            if (oSelectAll) { oSelectAll.setSelected(bAllSelected); }
        },

        // ─────────────────────────────────────────────────────────────────────
        // onSortFilterCA — Sort and Filter for CA table
        // ─────────────────────────────────────────────────────────────────────
        onSortFilterCA: function (oEvent) {
            var oTable = this.getView().byId("tableISUCA");
            var oBinding = oTable.getBinding("items");
            if (!oBinding) { return; }

            if (!this._oCASettingsDialog) {
                this._oCASettingsDialog = new sap.m.ViewSettingsDialog({
                    title: "Sort & Filter Contract Accounts",
                    sortItems: [
                        new sap.m.ViewSettingsItem({ key: "selected", text: "Selected First" }),
                        new sap.m.ViewSettingsItem({ key: "contacc", text: "Contract Account" }),
                        new sap.m.ViewSettingsItem({ key: "buspartner", text: "Business Partner" }),
                        new sap.m.ViewSettingsItem({ key: "est12month", text: "Est. Usage (MWh)" }),
                        new sap.m.ViewSettingsItem({ key: "metered", text: "Metered" }),
                        new sap.m.ViewSettingsItem({ key: "arrear60", text: "60-Day Arrear" }),
                        new sap.m.ViewSettingsItem({ key: "nsfFlag", text: "2-NSF" }),
                        new sap.m.ViewSettingsItem({ key: "replacementCA", text: "Replacement CA" }),
                        new sap.m.ViewSettingsItem({ key: "inactive", text: "Inactive" }),
                        new sap.m.ViewSettingsItem({ key: "cycle20", text: "Cycle 20" }),
                        new sap.m.ViewSettingsItem({ key: "providercontractId", text: "PC ID" }),
                        new sap.m.ViewSettingsItem({ key: "enrolled", text: "Enrolled" }),
                        new sap.m.ViewSettingsItem({ key: "validity", text: "Validity" })
                    ],
                    filterItems: [
                        new sap.m.ViewSettingsFilterItem({
                            key: "selected", text: "Selected",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "selected:true", text: "Selected" }),
                                new sap.m.ViewSettingsItem({ key: "selected:false", text: "Not Selected" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "contacc", text: "Contract Account",
                            items: [
                                // populated dynamically
                                // Note: static values won't scale — see note below
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "buspartner", text: "Business Partner",
                            items: [
                                // populated dynamically like contacc below
                            ]
                        }),

                        new sap.m.ViewSettingsFilterItem({
                            key: "metered", text: "Metered",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "metered:Yes", text: "Yes" }),
                                new sap.m.ViewSettingsItem({ key: "metered:No", text: "No" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "arrear60", text: "60-Day Arrear",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "arrear60:Yes", text: "Yes" }),
                                new sap.m.ViewSettingsItem({ key: "arrear60:No", text: "No" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "nsfFlag", text: "2-NSF",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "nsfFlag:Yes", text: "Yes" }),
                                new sap.m.ViewSettingsItem({ key: "nsfFlag:No", text: "No" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "replacementCA", text: "Replacement CA",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "replacementCA:Yes", text: "Yes" }),
                                new sap.m.ViewSettingsItem({ key: "replacementCA:No", text: "No" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "inactive", text: "Inactive",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "inactive:Yes", text: "Yes" }),
                                new sap.m.ViewSettingsItem({ key: "inactive:No", text: "No" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "cycle20", text: "Cycle 20",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "cycle20:Yes", text: "Yes" }),
                                new sap.m.ViewSettingsItem({ key: "cycle20:No", text: "No" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "providercontractId", text: "PC ID",
                            items: []   // populated dynamically
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "enrolled", text: "Enrolled",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "enrolled:Yes", text: "Yes" }),
                                new sap.m.ViewSettingsItem({ key: "enrolled:No", text: "No" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "validity", text: "Validity",
                            items: []   // populated dynamically
                        })
                    ],
                    confirm: function (oConfirmEvent) {
                        var mParams = oConfirmEvent.getParameters();

                        if (mParams.sortItem) {
                            var sSortKey = mParams.sortItem.getKey();
                            var bDesc = mParams.sortDescending;

                            if (sSortKey === "selected") {
                                oBinding.sort(new sap.ui.model.Sorter("selected", bDesc,
                                    false,
                                    function (a, b) {
                                        var valA = (a === true || a === "true") ? 1 : 0;
                                        var valB = (b === true || b === "true") ? 1 : 0;
                                        return valB - valA;
                                    }
                                ));
                            } else {
                                oBinding.sort(new sap.ui.model.Sorter(sSortKey, bDesc));
                            }
                        } else {
                            oBinding.sort(null);
                        }

                        if (mParams.filterItems && mParams.filterItems.length > 0) {
                            // var aFilters = mParams.filterItems.map(function (oItem) {
                            //     var aParts = oItem.getKey().split(":");
                            //     return new sap.ui.model.Filter(aParts[0], sap.ui.model.FilterOperator.EQ, aParts[1]);
                            // });
                            var aFilters = mParams.filterItems.map(function (oItem) {
                                var aParts = oItem.getKey().split(":");
                                var sField = aParts[0];
                                var sVal = aParts[1];

                                // Fields that store "" instead of "No"
                                var aEmptyMeansNo = ["arrear60", "nsfFlag", "replacementCA", "inactive", "cycle20", "enrolled", "providercontractId"];
                                if (aEmptyMeansNo.indexOf(sField) !== -1 && sVal === "No") {
                                    return new sap.ui.model.Filter(sField, sap.ui.model.FilterOperator.EQ, "");
                                }
                                // selected is stored as boolean, not string
                                if (sField === "selected") {
                                    return new sap.ui.model.Filter(sField, sap.ui.model.FilterOperator.EQ, sVal === "true");
                                }
                                return new sap.ui.model.Filter(sField, sap.ui.model.FilterOperator.EQ, sVal);
                            });
                            oBinding.filter(aFilters, sap.ui.model.FilterType.Application);
                        } else {
                            oBinding.filter([], sap.ui.model.FilterType.Application);
                        }
                    },
                    reset: function () {
                        var oTable = this.getView().byId("tableISUCA");
                        var oBinding = oTable.getBinding("items");
                        oBinding.filter([], sap.ui.model.FilterType.Application);
                        oBinding.sort(null);
                        this._aAllCAContextData = null; // clear cache so next open re-reads all rows
                    }.bind(this)
                });
                this.getView().addDependent(this._oCASettingsDialog);
            }
            // Just before this._oCASettingsDialog.open():
            var oTable = this.getView().byId("tableISUCA");
            var oBinding = oTable.getBinding("items");

            // Cache all CA values from the full unfiltered list on first open
            if (!this._aAllCAContextData) {
                this._aAllCAContextData = oBinding.getCurrentContexts()
                    .map(ctx => ctx.getObject());
            }
            var aAllData = this._aAllCAContextData;

            var aDynamicFields = [
                { key: "contacc", label: "contacc" },
                { key: "buspartner", label: "buspartner" },
                { key: "providercontractId", label: "providercontractId" },
                { key: "validity", label: "validity" }
            ];

            aDynamicFields.forEach(function (field) {
                var aVals = [...new Set(
                    aAllData.map(row => row[field.key]).filter(Boolean)
                )].sort();

                var oFilterItem = this._oCASettingsDialog.getFilterItems()
                    .find(f => f.getKey() === field.key);
                if (oFilterItem) {
                    oFilterItem.destroyItems();
                    aVals.forEach(function (val) {
                        oFilterItem.addItem(new sap.m.ViewSettingsItem({
                            key: field.key + ":" + val, text: val
                        }));
                    });
                }
            }.bind(this));
            this._oCASettingsDialog.open();
        },

        // ─────────────────────────────────────────────────────────────────────────
        // onSortFilterIP — Sort & Filter for Involved Parties tab
        // ─────────────────────────────────────────────────────────────────────────
        onSortFilterIP: function () {
            var oTable = this.getView().byId("table");
            var oBinding = oTable.getBinding("items");
            if (!oBinding) { return; }

            if (!this._oIPSettingsDialog) {
                this._oIPSettingsDialog = new sap.m.ViewSettingsDialog({
                    title: "Sort & Filter Involved Parties",
                    sortItems: [
                        new sap.m.ViewSettingsItem({ key: "buspartner", text: "Business Partner" }),
                        new sap.m.ViewSettingsItem({ key: "role", text: "Role" })
                    ],
                    filterItems: [
                        new sap.m.ViewSettingsFilterItem({
                            key: "role", text: "Role",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "role:Account", text: "Account" }),
                                new sap.m.ViewSettingsItem({ key: "role:Employee Responsible", text: "Employee Responsible" }),
                                new sap.m.ViewSettingsItem({ key: "role:Sales Representative", text: "Sales Representative" })
                            ]
                        })
                    ],
                    confirm: function (oConfirmEvent) {
                        var mParams = oConfirmEvent.getParameters();
                        if (mParams.sortItem) {
                            oBinding.sort(new sap.ui.model.Sorter(mParams.sortItem.getKey(), mParams.sortDescending));
                        } else {
                            oBinding.sort(null);
                        }
                        if (mParams.filterItems && mParams.filterItems.length > 0) {
                            var aFilters = mParams.filterItems.map(function (oItem) {
                                var aParts = oItem.getKey().split(":");
                                return new sap.ui.model.Filter(aParts[0], sap.ui.model.FilterOperator.EQ, aParts[1]);
                            });
                            oBinding.filter(aFilters, sap.ui.model.FilterType.Application);
                        } else {
                            oBinding.filter([], sap.ui.model.FilterType.Application);
                        }
                    }
                });
                this.getView().addDependent(this._oIPSettingsDialog);
            }
            this._oIPSettingsDialog.open();
        },

        // ─────────────────────────────────────────────────────────────────────────
        // onSortFilterPC — Sort & Filter for Provider Contract tab
        // Applies to both tablesmall (SMB) and tablelarge (LCVP) via shared binding
        // ─────────────────────────────────────────────────────────────────────────
        onSortFilterPC: function () {
            // Determine which table is visible
            var oTableSMB = this.getView().byId("tablesmall");
            var oTableLCVP = this.getView().byId("tablelarge");
            var oTable = (oTableSMB && oTableSMB.getVisible()) ? oTableSMB : oTableLCVP;
            if (!oTable) { return; }
            var oBinding = oTable.getBinding("items");
            if (!oBinding) { return; }

            if (!this._oPCSettingsDialog) {
                this._oPCSettingsDialog = new sap.m.ViewSettingsDialog({
                    title: "Sort & Filter Provider Contracts",
                    sortItems: [
                        new sap.m.ViewSettingsItem({ key: "product", text: "Product" }),
                        new sap.m.ViewSettingsItem({ key: "portfolio", text: "Portfolio" }),
                        new sap.m.ViewSettingsItem({ key: "startDate", text: "Start Date" }),
                        new sap.m.ViewSettingsItem({ key: "endDate", text: "End Date" }),
                        new sap.m.ViewSettingsItem({ key: "subspercent", text: "Subscription %" }),
                        new sap.m.ViewSettingsItem({ key: "enrollChk", text: "Enroll" }),
                        new sap.m.ViewSettingsItem({ key: "exportChk", text: "Export" })
                    ],
                    filterItems: [
                        new sap.m.ViewSettingsFilterItem({
                            key: "enrollChk", text: "Enroll",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "enrollChk:true", text: "Checked" }),
                                new sap.m.ViewSettingsItem({ key: "enrollChk:false", text: "Unchecked" })
                            ]
                        }),
                        new sap.m.ViewSettingsFilterItem({
                            key: "exportChk", text: "Export",
                            items: [
                                new sap.m.ViewSettingsItem({ key: "exportChk:true", text: "Checked" }),
                                new sap.m.ViewSettingsItem({ key: "exportChk:false", text: "Unchecked" })
                            ]
                        })
                    ],
                    confirm: function (oConfirmEvent) {
                        var mParams = oConfirmEvent.getParameters();
                        if (mParams.sortItem) {
                            var sSortKey = mParams.sortItem.getKey();
                            var bDesc = mParams.sortDescending;
                            if (sSortKey === "enrollChk" || sSortKey === "exportChk") {
                                oBinding.sort(new sap.ui.model.Sorter(sSortKey, bDesc, false,
                                    function (a, b) {
                                        var vA = (a === true || a === "true") ? 1 : 0;
                                        var vB = (b === true || b === "true") ? 1 : 0;
                                        return vB - vA;
                                    }
                                ));
                            } else {
                                oBinding.sort(new sap.ui.model.Sorter(sSortKey, bDesc));
                            }
                        } else {
                            oBinding.sort(null);
                        }
                        if (mParams.filterItems && mParams.filterItems.length > 0) {
                            var aFilters = mParams.filterItems.map(function (oItem) {
                                var aParts = oItem.getKey().split(":");
                                var sVal = aParts[1];
                                // Boolean filter — match both true/false and "true"/"false"
                                if (sVal === "true" || sVal === "false") {
                                    return new sap.ui.model.Filter({
                                        filters: [
                                            new sap.ui.model.Filter(aParts[0], sap.ui.model.FilterOperator.EQ, sVal === "true"),
                                            new sap.ui.model.Filter(aParts[0], sap.ui.model.FilterOperator.EQ, sVal)
                                        ],
                                        and: false
                                    });
                                }
                                return new sap.ui.model.Filter(aParts[0], sap.ui.model.FilterOperator.EQ, sVal);
                            });
                            oBinding.filter(aFilters, sap.ui.model.FilterType.Application);
                        } else {
                            oBinding.filter([], sap.ui.model.FilterType.Application);
                        }
                        // Apply same sort/filter to the other table's binding too
                        var oOther = (oTable === oTableSMB) ? oTableLCVP : oTableSMB;
                        if (oOther) {
                            var oOtherBinding = oOther.getBinding("items");
                            if (oOtherBinding) {
                                if (mParams.sortItem) { oOtherBinding.sort(oBinding.aSorters || []); }
                                oOtherBinding.filter(oBinding.aApplicationFilters || [], sap.ui.model.FilterType.Application);
                            }
                        }
                    }
                });
                this.getView().addDependent(this._oPCSettingsDialog);
            }
            this._oPCSettingsDialog.open();
        },

        // ─────────────────────────────────────────────────────────────────────────
        // onSortFilterPR — Sort & Filter for Prospect tab
        // ─────────────────────────────────────────────────────────────────────────
        onSortFilterPR: function () {
            var oTable = this.getView().byId("table11");
            var oBinding = oTable.getBinding("items");
            if (!oBinding) { return; }

            if (!this._oPRSettingsDialog) {
                this._oPRSettingsDialog = new sap.m.ViewSettingsDialog({
                    title: "Sort & Filter Prospects",
                    sortItems: [
                        new sap.m.ViewSettingsItem({ key: "siteAddLoc", text: "Site Address Location" }),
                        new sap.m.ViewSettingsItem({ key: "projectedCon", text: "Projected Consumption" }),
                        new sap.m.ViewSettingsItem({ key: "year", text: "Year" })
                    ],
                    filterItems: [],   // No boolean flags on Prospect — add if needed
                    confirm: function (oConfirmEvent) {
                        var mParams = oConfirmEvent.getParameters();
                        if (mParams.sortItem) {
                            oBinding.sort(new sap.ui.model.Sorter(mParams.sortItem.getKey(), mParams.sortDescending));
                        } else {
                            oBinding.sort(null);
                        }
                        if (mParams.filterItems && mParams.filterItems.length > 0) {
                            var aFilters = mParams.filterItems.map(function (oItem) {
                                var aParts = oItem.getKey().split(":");
                                return new sap.ui.model.Filter(aParts[0], sap.ui.model.FilterOperator.EQ, aParts[1]);
                            });
                            oBinding.filter(aFilters, sap.ui.model.FilterType.Application);
                        } else {
                            oBinding.filter([], sap.ui.model.FilterType.Application);
                        }
                    }
                });
                this.getView().addDependent(this._oPRSettingsDialog);
            }
            this._oPRSettingsDialog.open();
        },

        onSaveCA: function () {
            // this.getView().byId("tablesaveCA").setVisible(true);
            // this.getView().byId("tableISUCA").setVisible(false);
            MessageToast.show("Use the Save button to persist your CA selections.");
        },

        // ─────────────────────────────────────────────────────────────────────
        // onExportFile — Export Provider Contract details to Excel / CSV
        // ─────────────────────────────────────────────────────────────────────
        onExportFile: function () {
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var oSCDetail = this.getView().getModel("oSCDetail");
            if (!oTabModel || !oSCDetail) {
                MessageBox.error("Models not initialized. Fetch the Opportunity first.");
                return;
            }

            var sSalesType = oSCDetail.getProperty("/salesType");
            if (sSalesType === "One-time Sale") {
                MessageBox.warning("Export is not available for One-time Sale (FDD Req#51).");
                return;
            }

            var aCA = (oTabModel.getProperty("/ConsumptionDetails") || []).filter(function (r) {
                return r.selected === true || r.selected === 'true';
            });
            if (aCA.length === 0) {
                MessageBox.warning("Please select at least one Contract Account before exporting.");
                return;
            }

            // var aPC = oTabModel.getProperty("/ProviderContracts") || [];
            var aAllPC = oTabModel.getProperty("/ProviderContracts") || [];
            var aPC = aAllPC.filter(function (r) {
                return r.exportChk === true || r.exportChk === 'true';
            });

            if (aPC.length === 0) {
                MessageBox.warning("Please check at least one Export row in Provider Contract before exporting.");
                return;
            }
            var sOppid = oSCDetail.getProperty("/Oppid") || '';
            var sNTE = oSCDetail.getProperty("/nte") || '';
            var sCategory = oSCDetail.getProperty("/oppCategory") || '';
            var sExportDate = new Date().toLocaleDateString();
            var bIsLCVP = (sCategory === "MIGP - LCVP" || sCategory === "MIGP - Dedicated Array");
            var sReferralCode = oSCDetail.getProperty("/enrollmentReferralCode") || '';
            var sEmailAddress = oSCDetail.getProperty("/enrollmentEmailId") || '';
            var aExportRows = [];

            // ── Helper: format date to MM/DD/YYYY ───────────────────────────
            var fnFormatDate = function (sDate) {
                if (!sDate || sDate.trim() === '') return '';
                var d = new Date(sDate);
                if (isNaN(d)) return '';
                var mm = String(d.getMonth() + 1).padStart(2, '0');
                var dd = String(d.getDate()).padStart(2, '0');
                var yyyy = d.getFullYear();
                return mm + '/' + dd + '/' + yyyy;
            };

            aCA.forEach(function (oCA) {
                // var aPCLoop = aPC.length > 0 ? aPC : [{}];
                // aPCLoop.forEach(function (oPC) {
                aPC.forEach(function (oPC) {

                    // ── Term calculation ─────────────────────────────────────
                    var termMonths = 0;
                    try {
                        var dS = new Date(oPC.startDate);
                        var dE = new Date(oPC.endDate);
                        if (!isNaN(dS) && !isNaN(dE) && dE > dS) {
                            termMonths = (dE.getFullYear() - dS.getFullYear()) * 12
                                + (dE.getMonth() - dS.getMonth());
                        }
                    } catch (e) { }
                    var termYr = Math.floor(termMonths / 12);
                    var termMo = termMonths % 12;

                    if (bIsLCVP) {
                        // ── LCVP: exactly 10 columns ─────────────────────────
                        aExportRows.push({
                            "Contract Account": oCA.contacc || '',
                            // "Contracted Price": sNTE,
                            "Contracted Price": parseFloat(parseFloat(sNTE || 0).toFixed(2)),
                            "Subscription Amount": '',
                            // "Subscription Percentage": oPC.subspercent || '',
                            "Subscription Percentage": parseInt(oPC.subspercent || 0, 10),
                            "Discount Code": '',
                            // "Term Year": termYr || '',
                            // "Term Month": termMo || '',
                            "Term Year": parseInt(termYr || 0, 10),
                            "Term Month": parseInt(termMo || 0, 10),
                            "LCVP Start Date": fnFormatDate(oPC.startDate),
                            "LCVP End Date": fnFormatDate(oPC.endDate),
                            "Portfolio ID": oPC.portfolio || ''
                        });
                    } else {
                        // ── SMB: exactly 9 columns ───────────────────────────
                        aExportRows.push({
                            "Contract Account": oCA.contacc || '',
                            "Program Name": oPC.product || '',
                            // "Subscription Amount": oPC.fixedPrice || '',
                            // "Subscription Percentage": oPC.subspercent || '',
                            "Subscription Amount": parseFloat(parseFloat(oPC.fixedPrice || 0).toFixed(2)),
                            "Subscription Percentage": parseInt(oPC.subspercent || 0, 10),
                            "Referral Code": sReferralCode,
                            "Email Address": sEmailAddress,
                            "MIGP Start Date": fnFormatDate(oPC.startDate),
                            "MIGP End Date": '',
                            "MIGP Change Date": '',
                            "Portfolio ID": oPC.portfolio || '' // new require ment 
                        });
                    }
                });
            });

            if (aExportRows.length === 0) { MessageBox.warning("No data to export."); return; }

            try {
                var XLSX = window.XLSX;
                if (!XLSX) { throw new Error("SheetJS not loaded"); }

                var ws = XLSX.utils.json_to_sheet(aExportRows);
                var wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "ProviderContracts");

                var aColKeys = Object.keys(aExportRows[0]);
                ws["!cols"] = aColKeys.map(function (k) {
                    var maxLen = aExportRows.reduce(function (acc, r) {
                        return Math.max(acc, String(r[k] || '').length);
                    }, k.length);
                    return { wch: Math.min(maxLen + 2, 40) };
                });

                var sFileName = (bIsLCVP ? "LCVP_" : "SMB_") + "Export_" + sOppid + "_" +
                    new Date().toISOString().slice(0, 10) + ".xlsx";
                XLSX.writeFile(wb, sFileName);

                // this._stampExportDate(aPC, sExportDate);
                this._stampExportDate(aAllPC, sExportDate);
                this._silentSaveExportDate();
                MessageToast.show("Exported " + aExportRows.length + " row(s) → " + sFileName);

            } catch (e) {
                console.warn("SheetJS unavailable, falling back to CSV:", e.message);
                this._downloadCSV(aExportRows, sOppid);
                this._stampExportDate(aPC, sExportDate);
                this._silentSaveExportDate();
            }
        },

        _stampExportDate: function (aAllPC, sDate) {
            if (!aAllPC || !aAllPC.length) { return; }
            aAllPC.forEach(function (r) {
                if (r.exportChk === true || r.exportChk === 'true') {
                    r.exportDate = sDate;
                }
            });
            this.getOwnerComponent().getModel("oOpportunityjmodel")
                .setProperty("/ProviderContracts", aAllPC);
        },

        _silentSaveExportDate: async function () {
            try {
                var oModel = this.getOwnerComponent().getModel();
                var oHdr = this.getView().getModel("oSCDetail");
                var oTabs = this.getOwnerComponent().getModel("oOpportunityjmodel");
                var aAllPC = oTabs.getProperty("/ProviderContracts") || [];

                // Send ALL rows with their exportChk + exportDate state
                // Backend will update each row individually based on product+portfolio key
                var aPC = aAllPC.map(function (r) {
                    return {
                        product: r.product || '',
                        portfolio: r.portfolio || '',
                        exportChk: (r.exportChk === true || r.exportChk === 'true') ? 'true' : 'false',
                        exportDate: (r.exportChk === true || r.exportChk === 'true') ? (r.exportDate || '') : ''
                    };
                });

                if (!aPC.length) { return; }

                var oAction = oModel.bindContext("/saveExportDate(...)");
                oAction.setParameter("bundle", JSON.stringify({
                    oppId: oHdr.getProperty("/Oppid"),
                    quoteId: oHdr.getProperty("/quoteId"),
                    objectStatus: oHdr.getProperty("/objectStatus"),
                    providerContracts: aPC
                }));
                await oAction.execute();
                console.log("exportDate silently saved for",
                    oHdr.getProperty("/objectStatus") === "QUOTE"
                        ? "Quote: " + oHdr.getProperty("/quoteId")
                        : "Opp: " + oHdr.getProperty("/Oppid")
                );
            } catch (e) {
                console.error("_silentSaveExportDate failed:", e.message);
            }
        },

        _downloadCSV: function (aRows, sOppid) {
            if (!aRows || aRows.length === 0) { return; }
            var aHeaders = Object.keys(aRows[0]);
            var sCsv = aHeaders.join(",") + "\n";
            aRows.forEach(function (r) {
                sCsv += aHeaders.map(function (h) {
                    var v = String(r[h] || '').replace(/"/g, '""');
                    return (v.indexOf(',') !== -1 || v.indexOf('"') !== -1) ? '"' + v + '"' : v;
                }).join(",") + "\n";
            });
            var oBlob = new Blob([sCsv], { type: "text/csv;charset=utf-8;" });
            var sUrl = URL.createObjectURL(oBlob);
            var oLink = document.createElement("a");
            oLink.setAttribute("href", sUrl);
            oLink.setAttribute("download",
                "ProviderContracts_" + sOppid + "_" + new Date().toISOString().slice(0, 10) + ".csv");
            document.body.appendChild(oLink);
            oLink.click();
            document.body.removeChild(oLink);
            URL.revokeObjectURL(sUrl);
            MessageToast.show("Exported as CSV (Excel library not loaded).");
        },

        // ─────────────────────────────────────────────────────────────────────
        // onAddnewProspectRow / 2 (FDD Req#36)
        // ─────────────────────────────────────────────────────────────────────
        onAddnewProspectRow: function () {
            var that = this;
            var dialog = new sap.m.Dialog({
                title: "Confirmation", type: "Message", state: "Information", titleAlignment: "Center",
                content: new sap.m.Text({ text: this.getView().getModel("i18n").getProperty("add_prospect_msg") }),
                beginButton: new sap.m.Button({ text: "Yes", press: function () { dialog.close(); that.onAddnewProspectRow2(); } }),
                endButton: new sap.m.Button({ text: "No", press: function () { dialog.close(); } }),
                afterClose: function () { dialog.destroy(); }
            });
            dialog.open();
        },

        onAddnewProspectRow2: function () {
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var aData = oTabModel.getProperty("/Prospects") || [];
            aData.push({ siteAddLoc: "", year: "", projectedCon: "" });
            oTabModel.setProperty("/Prospects", aData);
        },

        onDeleteProspectRow: function (oEvent) {
            var that = this;

            // Button is directly inside cells → one getParent() = ColumnListItem
            var oButton = oEvent.getSource();
            var oListItem = oButton.getParent();   // ← ONE parent only

            // Named model — must pass model name to getBindingContext
            var oContext = oListItem.getBindingContext("oOpportunityjmodel");
            if (!oContext) {
                sap.m.MessageToast.show("Could not identify row.");
                return;
            }

            // Path = "/Prospects/2" → index = 2
            var iIndex = parseInt(oContext.getPath().split("/").pop(), 10);
            if (isNaN(iIndex) || iIndex < 0) {
                sap.m.MessageToast.show("Could not identify row index.");
                return;
            }

            var dialog = new sap.m.Dialog({
                title: "Confirm Delete", type: "Message", state: "Warning", titleAlignment: "Center",
                content: new sap.m.Text({ text: "Are you sure you want to delete this prospect row?" }),
                beginButton: new sap.m.Button({
                    text: "Delete", type: "Reject",
                    press: function () {
                        dialog.close();
                        var oTabModel = that.getOwnerComponent().getModel("oOpportunityjmodel");
                        var aData = oTabModel.getProperty("/Prospects") || [];
                        aData.splice(iIndex, 1);
                        oTabModel.setProperty("/Prospects", aData);
                        sap.m.MessageToast.show("Prospect row deleted. Save to confirm.");
                    }
                }),
                endButton: new sap.m.Button({ text: "Cancel", press: function () { dialog.close(); } }),
                afterClose: function () { dialog.destroy(); }
            });
            dialog.open();
        },


        onContract: function () { sap.ui.core.UIComponent.getRouterFor(this).navTo("RouteView2"); },

        _trim: function (v) { return (v == null) ? "" : String(v).trim(); },
        _num: function (v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; },
        _iso: function (v) {
            if (!v) { return ""; }
            var d = new Date(v);
            if (isNaN(d)) { return String(v); }
            return d.toISOString().slice(0, 10);
        },
        // ─────────────────────────────────────────────────────────────────────
        // _loadSalesPhaseItems
        // Populates Sales Phase ComboBox based on oppCategory
        // Only shows phases from current phase onwards
        // SMB  starts at: Approval
        // LCVP starts at: Structuring
        // ─────────────────────────────────────────────────────────────────────
        _loadSalesPhaseItems: function (sCategory, sCurrentPhase) {
            // ADD THIS TEMPORARILY
            console.log('_loadSalesPhaseItems → sCategory:', sCategory,
                '| sCurrentPhase:', sCurrentPhase,
                '| aAllPhases:', JSON.stringify(this._oSalesPhases[sCategory] || []));

            var oCB = this.getView().byId("idsalesPhase");
            if (!oCB) { return; }

            var bIsLCVP = (sCategory === "MIGP - LCVP" || sCategory === "MIGP - Dedicated Array");
            var bIsSMB = (sCategory === "MIGP - Small Business");
            var aAllPhases = (this._oSalesPhases[sCategory] || []);

            if (aAllPhases.length === 0) {
                oCB.destroyItems();
                oCB.setEnabled(false);
                return;
            }

            // ── Rebuild ComboBox ──────────────────────────────────────────────────
            oCB.destroyItems();

            // ── SMB: ALL phases come FROM C4C — fully read only ───────────────────
            // BTP only displays — user CANNOT change any SMB phase
            if (bIsSMB) {
                var oSMBPhase = aAllPhases.find(function (p) {
                    return p.code === sCurrentPhase || p.text === sCurrentPhase;
                });
                if (oSMBPhase) {
                    oCB.addItem(new sap.ui.core.Item({
                        key: oSMBPhase.code,
                        text: oSMBPhase.text + ' (set by C4C)'
                    }));
                    oCB.setSelectedKey(oSMBPhase.code);
                }
                oCB.setEnabled(false);
                console.log('SMB Sales Phase', sCurrentPhase,
                    '— all phases received from C4C');
                return;
            }

            // ── LCVP only beyond this point ───────────────────────────────────────

            // Phases that come FROM C4C in LCVP:
            // Z8 = Structuring  ← C4C pushes to BTP
            // Z3 = Final Review ← C4C pushes to BTP
            var aReadOnlyCodes = ['Z8', 'Z3'];

            // Find index of current phase
            var iIdx = aAllPhases.findIndex(function (p) {
                return p.code === sCurrentPhase || p.text === sCurrentPhase;
            });
            if (iIdx === -1) { iIdx = 0; }

            // Check if current phase is read-only (received from C4C)
            var bCurrentIsReadOnly = aReadOnlyCodes.indexOf(sCurrentPhase) !== -1;

            if (bCurrentIsReadOnly) {
                // ── Current phase came FROM C4C ───────────────────────────────────
                // Show current phase as display item (set by C4C)
                // BUT also show next available phases so user can move forward
                var oCurrentPhase = aAllPhases.find(function (p) {
                    return p.code === sCurrentPhase;
                });

                // Add current phase as first item (display only label)
                if (oCurrentPhase) {
                    oCB.addItem(new sap.ui.core.Item({
                        key: oCurrentPhase.code,
                        text: oCurrentPhase.text + ' (set by C4C)'
                    }));
                }

                // Add next available phases after current
                // User can move FORWARD from read-only phase
                aAllPhases.slice(iIdx + 1).forEach(function (oPhase) {
                    // Skip other read-only phases (Z3)
                    if (aReadOnlyCodes.indexOf(oPhase.code) !== -1) { return; }
                    // Skip Z4 Closure — only show when received from C4C
                    if (oPhase.code === 'Z4') { return; }

                    oCB.addItem(new sap.ui.core.Item({
                        key: oPhase.code,
                        text: oPhase.text
                    }));
                });

                // Enable so user can select next phase
                oCB.setEnabled(true);
                // Set current read-only phase as selected
                if (oCurrentPhase) {
                    oCB.setSelectedKey(oCurrentPhase.code);
                }

                console.log('LCVP Phase', sCurrentPhase,
                    '— set by C4C, next phases available');
                return;
            }

            // ── Current phase is user-selectable ─────────────────────────────────
            // Get phases from current index onwards
            var aAvailable = aAllPhases.slice(iIdx).filter(function (p) {
                // Remove read-only phases from dropdown
                if (aReadOnlyCodes.indexOf(p.code) !== -1) { return false; }
                // Remove Z4 unless already received from C4C
                if (p.code === 'Z4' &&
                    sCurrentPhase !== 'Z4' &&
                    sCurrentPhase !== 'Closure') {
                    return false;
                }
                return true;
            });

            aAvailable.forEach(function (oPhase) {
                oCB.addItem(new sap.ui.core.Item({
                    key: oPhase.code,
                    text: oPhase.text
                }));
            });

            oCB.setEnabled(true);

            // Set current phase as selected
            if (sCurrentPhase) {
                var oMatch = aAllPhases.find(function (p) {
                    return p.code === sCurrentPhase || p.text === sCurrentPhase;
                });
                if (oMatch) { oCB.setSelectedKey(oMatch.code); }
            }

            console.log("LCVP Sales Phase loaded | current:", sCurrentPhase,
                "| available:", aAvailable.map(function (p) {
                    return p.code + ':' + p.text;
                }));
        },

        // ─────────────────────────────────────────────────────────────────────
        // onSalesPhaseChange
        // Validates before allowing phase change:
        // LCVP → must have Sub% on ALL PC rows before moving to Approval
        // After validation → updates model + pushes to C4C via CPI
        // ─────────────────────────────────────────────────────────────────────
        // ─────────────────────────────────────────────────────────────────────
        // onSalesPhaseChange
        // Only updates the model — validations and CPI push happen on Save
        // ─────────────────────────────────────────────────────────────────────
        onSalesPhaseChange: function (oEvent) {
            var sNewCode = oEvent.getParameter("selectedItem").getKey();
            var oSCDetail = this.getView().getModel("oSCDetail");
            var sCategory = oSCDetail.getProperty("/oppCategory") || '';
            var sPrevPhase = oSCDetail.getProperty("/salesPhase") || '';

            // ── SMB — ComboBox is disabled, should never reach here ──────────
            if (sCategory === "MIGP - Small Business") {
                MessageBox.error(
                    "Sales Phase for Small Business is managed " +
                    "by SAP Sales Cloud. No changes allowed in BTP."
                );
                this.getView().byId("idsalesPhase").setSelectedKey(sPrevPhase);
                return;
            }

            // ── LCVP — just update model, Save handles validation + CPI push ─
            oSCDetail.setProperty("/salesPhase", sNewCode);
        },

        // ─────────────────────────────────────────────────────────────────────
        // onEnroll
        // Enroll button handler in Provider Contract tab
        // Validates:
        //   1. Enrollment Referral Code at header
        //   2. Enrollment Email ID at header
        //   3. At least one PC row has enrollChk checked
        //   4. At least one CA row is selected
        // Then calls initiateEnrollment action
        // ─────────────────────────────────────────────────────────────────────
        onEnroll: async function () {
            var oSCDetail = this.getView().getModel("oSCDetail");
            var oTabModel = this.getOwnerComponent().getModel("oOpportunityjmodel");
            var oModel = this.getOwnerComponent().getModel();

            // ── Step 1: Validate Referral Code ────────────────────────────────
            var sRefCode = oSCDetail.getProperty("/enrollmentReferralCode") || '';
            if (!sRefCode.trim()) {
                MessageBox.error(
                    "Enrollment Referral Code is required before enrolling."
                );
                return;
            }

            // ── Step 2: Validate Email ID ─────────────────────────────────────
            var sEmail = oSCDetail.getProperty("/enrollmentEmailId") || '';
            if (!sEmail.trim()) {
                MessageBox.error(
                    "Enrollment Email ID is required before enrolling."
                );
                return;
            }

            // ── Step 3: Validate checked PC rows ──────────────────────────────
            var aPCRows = oTabModel.getProperty("/ProviderContracts") || [];
            var aChecked = aPCRows.filter(function (r) {
                return r.enrollChk === true || r.enrollChk === 'true';
            });

            if (aChecked.length === 0) {
                MessageBox.error(
                    "Please check at least one Provider Contract row to enroll."
                );
                return;
            }

            // ── Step 4: Validate selected CA rows ─────────────────────────────
            var aCAs = oTabModel.getProperty("/ConsumptionDetails") || [];
            var aSelected = aCAs.filter(function (r) {
                return r.selected === true || r.selected === 'true';
            });

            if (aSelected.length === 0) {
                MessageBox.error(
                    "Please select at least one Contract Account before enrolling."
                );
                return;
            }

            // ── Step 5: Confirm before enrolling ──────────────────────────────
            var that = this;
            MessageBox.confirm(
                "Initiate enrollment for " + aChecked.length +
                " product row(s) and " + aSelected.length +
                " Contract Account(s)?\n\n" +
                "Referral Code: " + sRefCode + "\n" +
                "Email ID: " + sEmail,
                {
                    title: "Confirm Enrollment",
                    actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                    onClose: async function (sAction) {
                        if (sAction !== MessageBox.Action.YES) { return; }

                        // ── Step 6: Call initiateEnrollment ───────────────────
                        try {
                            that.getView().setBusy(true);

                            var oOp = oModel.bindContext(
                                "/initiateEnrollment(...)"
                            );
                            oOp.setParameter("bundle", JSON.stringify({
                                oppId: oSCDetail.getProperty("/Oppid"),
                                quoteId: oSCDetail.getProperty("/quoteId"),
                                objectStatus: oSCDetail.getProperty("/objectStatus"),
                                referralCode: sRefCode,
                                emailId: sEmail,
                                providerContracts: aChecked,
                                contractAccounts: aSelected
                            }));
                            await oOp.execute();

                            var oResult = oOp.getBoundContext().getObject();
                            var res = typeof oResult.value === "string"
                                ? JSON.parse(oResult.value)
                                : oResult.value;

                            that.getView().setBusy(false);

                            if (res && res.value === "ok") {
                                MessageBox.success(
                                    "Enrollment initiated successfully.\n" +
                                    (res.providerContractId
                                        ? "Provider Contract#: " +
                                        res.providerContractId
                                        : "Provider Contract# will be " +
                                        "returned by CR&B.")
                                );
                            } else {
                                MessageBox.error(
                                    "Enrollment failed: " +
                                    (res && res.message
                                        ? res.message
                                        : "Unknown error")
                                );
                            }

                        } catch (e) {
                            that.getView().setBusy(false);
                            MessageBox.error(
                                "Enrollment failed: " + e.message
                            );
                        }
                    }
                }
            );
        },
        // ─────────────────────────────────────────────────────────────────────
        // onSaveopp — full save (header + all 4 tabs)
        // ─────────────────────────────────────────────────────────────────────
        // onSaveopp — full save (header + all 4 tabs)
        // Validates LCVP sales phase, saves to BTP DB,
        // then pushes phase to C4C via CPI for BTP→C4C phases only
        // ─────────────────────────────────────────────────────────────────────
        onSaveopp: async function () {
            try {
                var oView = this.getView();
                var s = function (v) { return (v == null ? '' : String(v)); };
                var oModel = this.getOwnerComponent().getModel();
                var oHdr = oView.getModel("oSCDetail");
                var oTabs = this.getOwnerComponent().getModel("oOpportunityjmodel");

                if (!oHdr || !oTabs) {
                    MessageBox.error("Models not initialized. Fetch the Opportunity first.");
                    return;
                }

                // ── Quote context check ──────────────────────────────────────
                // var bIsQuote = oHdr.getProperty("/isQuoteContext") || false;
                // var bIsQuoteCtx = oHdr.getProperty("/isQuoteContext") || false;
                // ── Block save if Opportunity is Won and not in Quote context ────
                var sCurrentStatus = s(oHdr.getProperty("/status"));
                var bIsQuote = oHdr.getProperty("/isQuoteContext") || false;
                if (sCurrentStatus === "Won" && !bIsQuote) {
                    MessageBox.error(
                        "This Opportunity is Won and locked. " +
                        "Open the Quote to make amendments."
                    );
                    return;
                }
                var sOppid = this._trim(
                    oView.byId("il8").getValue() || oHdr.getProperty("/Oppid")
                );


                if (!sOppid) {
                    MessageBox.error("Please enter/fetch an Opportunity ID before saving.");
                    return;
                }

                if (!oHdr.getProperty("/objectStatus")) {
                    oHdr.setProperty("/objectStatus", bIsQuote ? 'QUOTE' : 'OPP');
                }
                oHdr.setProperty("/Oppid", sOppid);

                var s = this._trim.bind(this);
                var n = this._num.bind(this);
                var iso = this._iso.bind(this);

                var oppCategory = s(oView.byId("idModel_CS").getSelectedKey() || oHdr.getProperty("/oppCategory"));
                var priceStructure = s(oView.byId("idMode_CS").getSelectedKey() || oHdr.getProperty("/priceStructure"));
                var salesType = s(oView.byId("idMode_S").getSelectedKey() || oHdr.getProperty("/salesType"));
                var sStatus = s(oHdr.getProperty("/status"));
                var sSalesPhase = s(oHdr.getProperty("/salesPhase"));

                // ── LCVP Sales Phase validations ─────────────────────────────
                // Runs only for LCVP + only for BTP→C4C phases (Z0, Z9, Z4)
                if (oppCategory === "MIGP - LCVP" || oppCategory === "MIGP - Dedicated Array") {

                    // Z0 Approval — Sub% must be filled on ALL PC rows
                    if (sSalesPhase === "Z0") {
                        var aPCV = oTabs.getProperty("/ProviderContracts") || [];
                        if (aPCV.length === 0) {
                            MessageBox.error(
                                "No Provider Contract rows found. " +
                                "Please add product details before Approval."
                            );
                            return;
                        }
                        var bSubOk = aPCV.every(function (r) {
                            return String(r.subspercent || "").trim() !== "" &&
                                parseFloat(r.subspercent) > 0;
                        });
                        if (!bSubOk) {
                            MessageBox.error(
                                "Maintain Subscription % for all Provider Contract rows before Approval."
                            );
                            return;
                        }
                    }

                    // Z9 Contracting — Term + Start/End Dates must be filled
                    if (sSalesPhase === "Z9") {
                        var sTrm = s(oHdr.getProperty("/term"));
                        if (!sTrm || sTrm === "0" || sTrm === "0.00") {
                            MessageBox.error(
                                "Term must be filled before saving Contracting phase. " +
                                "Please run Calculate first."
                            );
                            return;
                        }
                        var aPCDates = oTabs.getProperty("/ProviderContracts") || [];
                        var bDatesOk = aPCDates.every(function (r) {
                            return r.startDate && r.startDate.trim() !== "" &&
                                r.endDate && r.endDate.trim() !== "";
                        });
                        if (!bDatesOk) {
                            MessageBox.error(
                                "Start Date and End Date must be filled " +
                                "for all Provider Contract rows before Contracting."
                            );
                            return;
                        }
                    }

                    // Z4 Closure — only allowed after BTP has received Z3 from C4C
                    // _loadSalesPhaseItems already blocks Z4 unless prevPhase = Z3
                    // This is a safety net in case model is in unexpected state
                    if (sSalesPhase === "Z4") {
                        console.log("Saving Closure phase — triggered after Final Review received from C4C.");
                    }
                }

                // ── Build header ─────────────────────────────────────────────
                var header = {
                    Oppid: sOppid,
                    quoteId: s(oHdr.getProperty("/quoteId")),
                    objectStatus: s(oHdr.getProperty("/objectStatus")),
                    oppCategory: oppCategory,
                    status: sStatus,
                    annualGross: String(n(oHdr.getProperty("/annualGross"))),
                    sizeofOpp: String(n(oHdr.getProperty("/sizeofOpp"))),
                    percentUsage: String(n(oHdr.getProperty("/percentUsage"))),
                    annualSubs: String(n(oHdr.getProperty("/annualSubs"))),
                    oppSubs: String(n(oHdr.getProperty("/oppSubs"))),
                    contLifetimeval: String(n(oHdr.getProperty("/contLifetimeval"))),
                    salesType: salesType,
                    priceStructure: priceStructure,
                    nte: String(n(oHdr.getProperty("/nte"))),
                    term: String(n(oHdr.getProperty("/term"))),
                    meteredCon: String(n(oHdr.getProperty("/meteredCon"))),
                    unmeteredCon: String(n(oHdr.getProperty("/unmeteredCon"))),
                    commencementLetterSent: String(oHdr.getProperty("/commencementLetterSent") || false),
                    commencementLetterSigned: String(oHdr.getProperty("/commencementLetterSigned") || false),
                    enrolled: s(oHdr.getProperty("/enrolled")),
                    includeInPipeline: String(oHdr.getProperty("/includeInPipeline") || false),
                    salesPhase: sSalesPhase,
                    enrollmentReferralCode: s(oHdr.getProperty("/enrollmentReferralCode")),
                    enrollmentEmailId: s(oHdr.getProperty("/enrollmentEmailId")),
                    // lastFetchedAt: s(oHdr.getProperty("/lastFetchedAt")),   // ← persisted timestamp from DB
                    lastFetchedAt: s(oHdr.getProperty("/lastFetchedAtISO")),
                };

                // ── Build tab data ───────────────────────────────────────────
                var involvedParties = (oTabs.getProperty("/InvolvedParties") || []).map(function (r) {
                    return { buspartner: s(r.buspartner), role: s(r.role) };
                });

                var providerContracts = (oTabs.getProperty("/ProviderContracts") || []).map(function (r) {
                    return {
                        product: s(r.product),
                        portfolio: s(r.portfolio),
                        fixedMWh: s(r.fixedMWh),
                        fixedPrice: s(r.fixedPrice),
                        netPremium: s(r.netPremium),
                        portfolioprice: s(r.portfolioprice),
                        subspercent: s(r.subspercent),
                        startDate: s(iso(r.startDate)),
                        endDate: s(iso(r.endDate)),
                        estUsage: s(r.estUsage),
                        recPrice: s(r.recPrice),
                        migpPercent: s(r.migpPercent),
                        recpercent: s(r.recpercent),
                        migpMWh: s(r.migpMWh),
                        recMWh: s(r.recMWh),
                        enrollChk: r.enrollChk === true || r.enrollChk === 'true' ? 'true' : 'false',
                        exportChk: r.exportChk === true || r.exportChk === 'true' ? 'true' : 'false',
                        exportDate: s(r.exportDate)
                    };
                });

                var consumptionDetails = (oTabs.getProperty("/ConsumptionDetails") || []).map(function (r) {
                    return {
                        contacc: s(r.contacc),
                        buspartner: s(r.buspartner),
                        validity: s(iso(r.validity)),
                        replacementCA: s(r.replacementCA),
                        est12month: s(r.est12month),
                        est12monthMetered: s(r.est12monthMetered),
                        est12monthUnMetered: s(r.est12monthUnMetered),
                        RECBlender_flag: s(r.RECBlender_flag),
                        providercontractId: s(r.providercontractId),
                        cycle20: s(r.cycle20),
                        metered: s(r.metered),
                        arrear60: s(r.arrear60),
                        nsfFlag: s(r.nsfFlag),
                        inactive: s(r.inactive),
                        enrolled: s(r.enrolled),
                        selected: r.selected === true || r.selected === 'true' ? 'true' : 'false'
                    };
                });

                var prospects = (oTabs.getProperty("/Prospects") || []).map(function (r) {
                    return {
                        siteAddLoc: s(r.siteAddLoc),
                        projectedCon: s(r.projectedCon),
                        year: s(r.year)
                    };
                });

                console.log('onSaveopp bundle: IP=' + involvedParties.length +
                    ' PC=' + providerContracts.length +
                    ' CD=' + consumptionDetails.length +
                    ' PR=' + prospects.length +
                    ' isQuote=' + bIsQuote +
                    ' salesPhase=' + sSalesPhase);

                // ── Save to BTP DB ───────────────────────────────────────────
                var oAction = oModel.bindContext("/saveFullOpportunity(...)");
                oAction.setParameter("bundle", JSON.stringify({
                    header, involvedParties, providerContracts, consumptionDetails, prospects
                }));
                await oAction.execute();

                var oRes = oAction.getBoundContext().getObject();
                var res = (typeof oRes.value === "string") ? JSON.parse(oRes.value) : oRes.value;

                if (res && res.value === "ok") {
                    MessageBox.success(
                        bIsQuote
                            ? "Quote and all tab data saved successfully."
                            : "Opportunity and all tab data saved successfully."
                    );
                    // this.oncalc();  // ← ADD THIS — recalculates header fields after save
                    console.log('CPI push check — oppCategory:', oppCategory, '| salesPhase:', sSalesPhase);
                    // ── CPI Push — LCVP BTP→C4C phases only ─────────────────
                    // Z0 = Approval    → BTP pushes to C4C
                    // Z9 = Contracting → BTP pushes to C4C
                    // Z4 = Closure     → BTP pushes to C4C (after receiving Z3)
                    // Z8, Z3 are C4C→BTP — BTP never pushes these back
                    var aBTPPushPhases = ["Z0", "Z9", "Z4"];
                    if ((oppCategory === "MIGP - LCVP" || oppCategory === "MIGP - Dedicated Array") &&
                        aBTPPushPhases.indexOf(sSalesPhase) !== -1) {
                        try {
                            var oPhaseOp = oModel.bindContext("/pushSalesPhaseToC4C(...)");
                            oPhaseOp.setParameter("bundle", JSON.stringify({
                                oppId: sOppid,
                                quoteId: s(oHdr.getProperty("/quoteId")),           // ← ADD
                                objectStatus: s(oHdr.getProperty("/objectStatus")), // ← ADD
                                salesPhase: sSalesPhase,
                                enrollmentReferralCode: s(oHdr.getProperty("/enrollmentReferralCode")),   // ← add
                                enrollmentEmailId: s(oHdr.getProperty("/enrollmentEmailId")),              // ← add
                                salesPhaseText: "",   // backend resolves via oPhaseMap
                                status: sStatus,
                                oppCategory: oppCategory
                            }));
                            // await oPhaseOp.execute();
                            // MessageToast.show("Sales Phase synced to SAP Sales Cloud.");
                            await oPhaseOp.execute();
                            var oPhaseRes = oPhaseOp.getBoundContext().getObject();
                            var phaseRes = (typeof oPhaseRes.value === "string")
                                ? JSON.parse(oPhaseRes.value)
                                : oPhaseRes.value;

                            if (phaseRes && phaseRes.value === "ok") {
                                MessageToast.show("Sales Phase synced to SAP Sales Cloud.");
                            } else {
                                // Show the actual error from backend so you can diagnose
                                var sErrMsg = (phaseRes && phaseRes.message)
                                    ? phaseRes.message
                                    : "C4C sync pending.";
                                MessageBox.warning("Saved to BTP.\n\n" + sErrMsg);
                            }
                            // } catch (ePush) {
                            //     console.error("pushSalesPhaseToC4C failed:", ePush.message);
                            //     MessageToast.show("Saved to BTP. C4C sync pending.");
                            // }
                        } catch (ePush) {
                            console.error("pushSalesPhaseToC4C failed:", ePush.message);
                            MessageToast.show("C4C sync failed: " + ePush.message); // ← show real error
                        }

                    }

                } else {
                    MessageBox.error("Save failed: " + (res && res.message ? res.message : "Unknown error"));
                }

            } catch (e) {
                MessageBox.error("Save failed: " + e.message);
            }
        },

        onSave: function () { console.log("onSave pressed"); }
    });
});