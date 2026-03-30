sap.ui.define([
    "sap/ui/core/mvc/Controller"
], (Controller) => {
    "use strict";

    return Controller.extend("com.sap.salescloud.controller.View2", {
        onInit() {

        },
        onComboBoxtype: function () {
            var b = this.getView().byId("idModelCS").getSelectedKey();
            if (b === "small") {
                this.getView().byId("tablesmall9").setVisible(true);
                this.getView().byId("tablelarge9").setVisible(false);
            }
            else if (b === "large") {
                this.getView().byId("tablelarge9").setVisible(true);
                this.getView().byId("tablesmall9").setVisible(false);
            }

        },

        onfetchCA: function () {
            this.getView().byId("tablesaveCA").setVisible(false);
            this.getView().byId("tableISUCA").setVisible(true);
        },
        onSaveCA: function () {
            this.getView().byId("tablesaveCA").setVisible(true);
            this.getView().byId("tableISUCA").setVisible(false);
        },
        onAddnewProspectRow: function () {
            var that = this;
            var dialog = new sap.m.Dialog({
                title: "Confirmation",
                type: 'Message',
                state: 'Information',
                titleAlignment: 'Center',
                content: new sap.m.Text({
                    text: this.getView().getModel("i18n").getProperty("add_prospect_msg")
                }),
                beginButton: new sap.m.Button({
                    text: "Yes",
                    press: function (oEvt) {
                        dialog.close();
                        that.onAddnewProspectRow2();
                    }
                }),
                endButton: new sap.m.Button({
                    text: "No",
                    press: function (oEvt) {
                        dialog.close();
                    }
                }),

                afterClose: function () {
                    dialog.destroy();
                }
            });
            dialog.open();
        },
        onAddnewProspectRow2: function (oEvt) {
            // var oSpecialsjmodel = this.getView().getModel("oSpecialsjmodel");
            // var aData = oSpecialsjmodel.getProperty("/fuses") || []; // Get current items
            this.sequenceCounter = 0;
            if (aData.length === 0) {
                this.sequenceCounter = 1;
            } else {
                this.sequenceCounter = aData.length;
                this.sequenceCounter = this.sequenceCounter + 1;
            }

            // Create a new item with empty values
            var newItem = {
                dummyBusiness: "",   // For the Input field
                dummyContract: "",  // For the ComboBox
                fuseVoltage: "",// For the ComboBox
                estUsage: ""
            };
            aData.push(newItem);
            // oSpecialsjmodel.setProperty("/fuses", aData);
        },

        onOpp: function () {
            var oRouter = sap.ui.core.UIComponent.getRouterFor(this);
            oRouter.navTo("RouteView1");
        }
        // onSave: function () {

        //     var oSCDetail1 = this.getView().getModel("oSCDetail");
        //     oSCDetail1.setData({
        //         cust_name: oSCDetail1.cust_name,
        //         emer_phone: oSCDetail1.emer_phone
        //     });
        // }
    });
});