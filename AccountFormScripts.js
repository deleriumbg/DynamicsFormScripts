var Sdk = window.Sdk || {};

(function () {
    // Code to run in the form OnSave event
    this.formOnSave = function (executionContext) {
        debugger;
        var formContext = executionContext.getFormContext();
        let customCreationDate = formContext.getAttribute("new_customcreationdate");
        if (customCreationDate.getValue() == null) {
            preventAutoSave(executionContext);
            Xrm.Navigation.openAlertDialog({ text: "Record can not be saved. Custom Creation Date field is required." });
            return;
        }

        let currentDate = new Date();
        let diffTime = currentDate.getTime() - customCreationDate.getValue().getTime();
        let diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        Xrm.WebApi.retrieveRecord("new_systemrules", "6FCCBE6E-699D-E911-A980-000D3A26C11D", "?$select=new_rulevalue").then(
            function success(result) {
                let systemRuleValue = result.new_rulevalue;
                if (systemRuleValue == null) {
                    preventAutoSave(executionContext);
                    Xrm.Navigation.openAlertDialog({ text: "Record can not be saved. System Rule value is required." });
                    return;
                }
                console.log("Retrieved Rule value: " + systemRuleValue);

                if (diffDays < systemRuleValue) {
                    preventAutoSave(executionContext);
                    Xrm.Navigation.openAlertDialog({ text: "Record can not be saved. Custom Creation Date doesn't meet the System Rule requirements." });
                } else {
                    // Define the data to create new account
                    const entityId = formContext.data.entity.getId().slice(1, -1);
                    var data =
                    {
                        "new_recordname": formContext.getAttribute("name").getValue() + "_" + currentDate.toLocaleDateString("en-US"),
                        "new_creationdate": customCreationDate.getValue(),
                        "new_deactivationdate": currentDate,
                        "new_deactivationperiod": Number(systemRuleValue),
                        "new_Account@odata.bind": "/accounts(" + entityId + ")"
                    }

                    // Create Account deactivation history record
                    Xrm.WebApi.createRecord("new_accountdeactivationhistory", data).then(
                    function success(result) {
                        console.log("Account deactivation history record created with ID: " + result.id);

                        // Set Account entity to Inactive status and set its Date to Deactivate to the current date
                        formContext.getAttribute("new_datetodeactivate").setValue(currentDate);
                        SetStateRequest(entityId); 
                    },
                    function (error) {
                        console.log(error.message);
                    });
                }  
            },
            function (error) {
                console.log(error.message);
            }
        );   
    }

    // Code to run in the VIP Account attribute OnChange event 
    this.vipAccountOnChange = function (executionContext) {
        debugger;
        var formContext = executionContext.getFormContext();
        let vipAccountNameObj = formContext.getAttribute("new_vipaccount_name");
        let tabObj = formContext.ui.tabs.get("SUMMARY_TAB");
        let sectionObj = tabObj.sections.get("VIP_Account_Details");

        if (formContext.getAttribute("new_vipaccount").getValue() == 1) {
            vipAccountNameObj.setValue(formContext.getAttribute("name").getValue());
            sectionObj.setVisible(true); //to show
        } else {
            vipAccountNameObj.setValue("");
            sectionObj.setVisible(false); //to hide
        }
    }
}).call(Sdk);

function preventAutoSave(executionContext) {  
    var eventArgs = executionContext.getEventArgs();  
    if (eventArgs.getSaveMode() == 70 || eventArgs.getSaveMode() == 2) {  
        eventArgs.preventDefault();  
    }  
} 

function SetStateRequest(entityId) {
    var entity = {};
    entity.statuscode = 2;
    entity.statecode = 1;

    var req = new XMLHttpRequest();
    req.open("PATCH", Xrm.Page.context.getClientUrl() + "/api/data/v8.1/accounts(" + entityId + ")", true);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
                Xrm.Navigation.openAlertDialog({ text: "Record successfully deactivated." });
            } else {
                Xrm.Navigation.openAlertDialog({ text: this.statusText });
            }
        }
    };
    req.send(JSON.stringify(entity));
}