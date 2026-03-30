sap.ui.define([
    "sap/ui/core/UIComponent",
    "com/sap/salescloud/model/models"
], (UIComponent, models) => {
    "use strict";

    return UIComponent.extend("com.sap.salescloud.Component", {
        metadata: {
            manifest: "json",
            interfaces: [
                "sap.ui.core.IAsyncContentCreation"
            ]
        },

        init() {
            // call the base component's init function
            UIComponent.prototype.init.apply(this, arguments);

            // set the device model
            this.setModel(models.createDeviceModel(), "device");

            // enable routing
            this.getRouter().initialize();

            //User Scope model
            var oUserScopeJModel = new sap.ui.model.json.JSONModel();
            this.setModel(oUserScopeJModel, "oUserScopeJModel");
        }
    });
});