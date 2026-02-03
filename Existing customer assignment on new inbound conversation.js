//Automatically identify and link existing Dynamics 365 customer records to Omnichannel inbound conversations using JavaScript and CE Web API
//Register method against onload event of conversation record.
function SetContact(executionContext) {
    try {
        window.top.Microsoft.Omnichannel.getConversationId().then(
 
            function success(ConversationId) {
 
                if (ConversationId != null) {
                    debugger;
                    checkIfContactAlreadyLinked(executionContext, ConversationId);
                }
            },
            function error(err) {
                debugger;
                console.log("Error in Getting Conversation ID. " + err);
            }
        );
    }
    catch (excep) {
        console.log(excep)
    }
}
 
 
function wait(ms) {
 
    var start = new Date().getTime();
 
    var end = start;
 
    while (end < start + ms) {
 
        end = new Date().getTime();
 
    }
 
    Xrm.Utility.closeProgressIndicator();
 
}
 
 
 //Checked is contact is already exist.
 
function checkIfContactAlreadyLinked(executionContext, ConversationId) {
 
    Xrm.WebApi.online.retrieveMultipleRecords("msdyn_ocliveworkitem", "?$select=_msdyn_customer_value&$filter=msdyn_ocliveworkitemid eq '" + ConversationId + "'").then(
 
        function success(data) {
            debugger;
            if ((data !== null) && (data.entities !== null) && (data.entities.length > 0) && data.entities[0]._msdyn_customer_value == null) {
 
                createContactLinkContact(ConversationId, executionContext);
 
            }
 
            else {
                debugger;
                console.log("Omnichannel: Contact Already Created and Linked ",);
 
            }
            debugger;
            console.log("2 GUID: ",);
 
        }
 
    );
 
}
 
 
//Link customer with the conversation 
function createContactLinkContact(ConversationId, executionContext) {
    debugger;
 
    var phone;
 
    var req = new XMLHttpRequest();
    req.open("GET", Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/msdyn_ocliveworkitems(" + ConversationId + ")?$select=subject", false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Prefer", "odata.include-annotations=*");
    req.send();
    if (req.status === 200) {
        var result = JSON.parse(req.response);
        console.log(result);
        // Columns
        var activityid = result["activityid"]; // Guid
        var subject = result["subject"]; // Text
        var _phone = subject.split(':');
        // phone = _phone[0];
        phone = _phone[0].replace('+', '');
        GetCustomerDetails(executionContext, ConversationId, phone);
    }
    else {
        console.log(req.responseText);
    }
}
 
 //get customer details in the base of phone number
function GetCustomerDetails(executionContext, conversationId, PhoneNumber) {
    debugger;
 
    var req = new XMLHttpRequest();
    req.open("GET", Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/contacts?$select=firstname,fullname,lastname&$filter=(contains(telephone1,'" + PhoneNumber + "') or contains(address1_telephone1,'" + PhoneNumber + "'))", false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Prefer", "odata.include-annotations=*,odata.maxpagesize=1");
    req.send();
 
    if (req.status === 200) {
        var result = JSON.parse(req.response);
 
        if (result.value && result.value.length > 0) {  // Ensure data exists
            var contact = result.value[0]; // Get first contact
            var contactid = contact.contactid; // Guid
            var fullname = contact.fullname; // Text
 
            Microsoft.Omnichannel.linkToConversation("contact", contactid, conversationId).then((response) => {
                // Refreshing the tab UI  
                Microsoft.Apm.getFocusedSession().getFocusedTab().refresh();
            }, (error) => {
                console.log(error);
            });
        } else {
            console.log("No matching contact found.");
        }
    } else {
        console.log(req.responseText);
    }
}
