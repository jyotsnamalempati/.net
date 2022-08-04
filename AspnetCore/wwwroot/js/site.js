// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

(function () {


    Office.initialize = function (reason) {
    //   $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        // Add a ViewSelectionChanged event handler.
        Office.context.document.addHandlerAsync(
            Office.EventType.ResourceSelectionChanged,
            getResourceGuid);

        Office.context.auth.getAccessTokenAsync(function (result) {
            if (result.status === "succeeded") {

                debugger;
                var token = result.value;
                console.log("token:" + token);
                // ...
            } else {
                console.log("Error obtaining token", result.error);
            }
        });


        $(document).ready(function () {
         
            showDocumentProperties();
           
        });
        
       
    


        function close() {
            window.open('', '_self').close();
        }

        //  }
        // If you need to initialize something you can do so here.
    };

})();

function pixelToPercentage(heightInPixel, widthInPixel) {
    var cheight = window.height ? window.height : screen.height;
    var cwidth = window.height ? window.width : screen.width;
    if (heightInPixel > cheight) heightInPixel = cheight;
    if (widthInPixel > cwidth) widthInPixel = cwidth;
    var height = Math.floor(100 * heightInPixel / cheight);
    var width = Math.floor(100 * widthInPixel / cwidth);
    return { height: height, width: width };
}


function showDocumentProperties() {
    var output = String.format(
        'The document mode is {0}.<br/>The URL of the active project is {1}.',
        Office.context.document.mode,
        Office.context.document.url);
    console.log(output);

    async function Gettoken() {

        try {
            let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
                allowSignInPrompt: true,
            });
            console.log(userTokenEncoded);
            
            $("#txtarea").val(userTokenEncoded);

        } catch (error) {
            console.log(error);
        }
    }

   

    
}

function Testdialog() {

    Office.context.auth.getAccessTokenAsync(function (result) {
        if (result.status === "succeeded") {
            var token = result.value;
            console.log("token:" + token);
            // ...
        } else {
            console.log("Error obtaining token", result.error);
        }
    });

    console.log("token");
    var tokenData = OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });



    console.log(tokenData);


}
async function LoginMsal() {
    const msalConfig = {
        auth: {
            clientId: "6c8f5ceb-ce5e-4646-8e4a-43bad62d265c",
            authority: "https://login.microsoftonline.com/common",
            knownAuthorities: [],
            redirectUri: "https://localhost:3001",
            postLogoutRedirectUri: "https://localhost:3001/logout",
            navigateToLoginRequestUrl: true,
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: false,
        }
    }


    // Create an instance of PublicClientApplication
    const msalInstance = new UserAgentApplication(msalConfig)

    // Handle the redirect flows
    msalInstance
        .handleRedirectPromise()
        .then((tokenResponse) => {
            console.log("tokenresponse");
        })
        .catch((error) => {
            console.log("error");
        });

}
function dialogCallback(asyncResult) {

    if (asyncResult.status == "failed") {

        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                showNotification("Domain is not trusted");
                break;
            case 12005:
                showNotification("HTTPS is required");
                break;
            case 12007:
                showNotification("A dialog is already opened.");
                break;
            default:
                showNotification(asyncResult.error.message);
                break;
        }
    }
    else {

        dialog = asyncResult.value;

        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);

    }
}

function closeModal() {
    Office.context.ui.messageParent("close");

}

function messageHandler(arg) {
    dialog.close();
}
async function getUserData() {
    try {
        console.log("getUserData");
        let userTokenEncoded = await OfficeRuntime.auth.getAccessToken();
        let userToken = jwt_decode(userTokenEncoded); // Using the https://www.npmjs.com/package/jwt-decode library.
        console.log(userToken.name); // user name
        console.log(userToken.preferred_username); // email
        console.log(userToken.oid); // user id     
    }
    catch (exception) {
        if (exception.code === 13003) {
            console.log(exception.code);
            // SSO is not supported for domain user accounts, only
            // Microsoft 365 Education or work account, or a Microsoft account.
        } else {
            console.log(exception.code);
            // Handle error
        }
    }
}

function AccessDocument() {
    
    localStorage.setItem("DocUrl", Office.context.document.url);

    // console.log("shareDocument()=>documenturl:" + documenturl);
    //Office.context.ui.displayDialogAsync(window.location.origin + "/ShareDocument",
    //    { height: 25, width: 40 }, dialogCallback);

    var dialog;
    Office.context.ui.displayDialogAsync(window.location.origin + "/ShareDocument", { height: 25, width: 40 },
        function (asyncResult) {
            dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, "Send for review");
        }
    );
}


function ClosePage() {
    window.open('', '_self').close();
}
function shareDocument() {
    var tokenvalue = "";
    
    
    localStorage.setItem("DocUrl", Office.context.document.url);

    let userTokenEncoded =      OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
    }).then(value => {

        localStorage.setItem("BTServiceToken", value);
        var dim = pixelToPercentage(360, 600);
        //Office.context.ui.displayDialogAsync(window.location.origin + "/ShareDocument",
        //    { height: 38, width: 45, displayInIframe: true, message: "Send for review" }, dialogCallback);

        Office.context.ui.displayDialogAsync(window.location.origin + "/ShareDocument",
            { height: dim.height, width: dim.width, displayInIframe: true, message: "Send for review" }, dialogCallback);
    }, reason => {
        console.log(reason);
        debugger;
    });
  
   
  
    
   

}

async function Gettoken() {

    try {
        let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
        });
        debugger;        
        $("#txtarea").val(userTokenEncoded);

    } catch (error) {
        console.log(error);
    }
}

function getUserList() {
    $("#toEmailSelection").select2({ width: 'resolve' });
    $("#toEmailSelection").prop("disabled", true);


    debugger;

    var token = localStorage.getItem("BTServiceToken");


    var data = [];
    $.ajax({
        url: "https://btserviceapiappservice.azurewebsites.net/api/User/getUsersFromAD",
        headers: {
            Authorization: 'Bearer ' + token
        },
        type: 'GET',
        success: function (result) {
            if (result?.status) {
                 data = result?.data
                    .map((item) => {
                        return { text: `${item.displayName} (${item.email})`, id: item.email };
                    });
                $("#toEmailSelection").select2("destroy");
                $("#toEmailSelection").select2({
                    width: 'resolve',
                    data: data
                })
            }
            $("#toEmailSelection").prop("disabled", false);
        },
        error: function (error) {
            console.log(error);
            $("#toEmailSelection").prop("disabled", false);
        }
    })
}





function getUserResponseList() {
    document.getElementById("overlay").style.display = "flex";

    var token = localStorage.getItem("BTServiceToken");


    $("#toEmailSelection").select2({ width: 'resolve' });
    $("#toEmailSelection").prop("disabled", true);
    var data = [];
    $.ajax({
        url: "https://btserviceapiappservice.azurewebsites.net/api/Document/getrespondemails",
        headers: {
            Authorization: 'Bearer ' + token
        },
        type: 'GET',
        data: { 'ItemPath': localStorage.getItem("ResDocUrl").split('Documents')[1]  },
        success: function (result) {
            if (result?.status) {
                var selectedArray = [];
                data = result?.data?.allUsers
                    .map((item) => {
                        selectedArray.push(item.email)
                        return { text: `${item.displayName} (${item.email})`, id: item.email };
                    });
                $("#toEmailSelection").select2("destroy");
                var selectInst = $("#toEmailSelection").select2({
                    width: 'resolve',
                    data: data
                })
                if (result?.data?.selectedUsers) {
                    selectedArray = result?.data?.selectedUsers;

                }
                selectInst.val(selectedArray).trigger("change");

            }
            $("#toEmailSelection").prop("disabled", false);
            document.getElementById("overlay").style.display = "none";
        },
        error: function (error) {
            console.log(error);
            $("#toEmailSelection").prop("disabled", false);
            document.getElementById("overlay").style.display = "none";
        }
    })
}


// modal open for respond
function showRespond() {


    var tokenvalue = "";
    localStorage.setItem("DocUrl", Office.context.document.url);

    let userTokenEncoded = OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
    }).then(value => {

        localStorage.setItem("BTServiceToken", value);
        var dim = pixelToPercentage(360, 600);
        //Office.context.ui.displayDialogAsync(window.location.origin + "/ShareDocument",
        //    { height: 38, width: 45, displayInIframe: true, message: "Send for review" }, dialogCallback);

        Office.context.ui.displayDialogAsync(window.location.origin + "/Respond",
            { height: dim.height, width: dim.width, displayInIframe: true, message: "Send for review" }, dialogCallback);
    }, reason => {
        console.log(reason);
        debugger;
    });    
}


// send for review
function SendDocumentforReview() {

    var path = localStorage.getItem("DocUrl");;
    var emailids = [];
    $(".select2-selection__choice__display").each(function (data) {
        emailids.push($(this).text().split("(")[1].split(")")[0]);
    });

    var token = localStorage.getItem("BTServiceToken");

    var msg = document.getElementById("txtmessage").value;

    let payLoad = JSON.stringify(
        {
            "itemPath": localStorage.getItem("DocUrl").split('Documents')[1],
            "emailIDs": emailids,
            "message": msg,
            "source": "AddIns",
            "documentActivityMasterID": 2
        });
    document.getElementById("overlay").style.display = "flex";
    $.ajax({
        url: "https://btserviceapiappservice.azurewebsites.net/api/Document/sendForReviewAddIns",
        headers: {
            'Authorization': 'Bearer ' + token,
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        },
        type: 'POST',
        data: payLoad,
        dataType: 'json',
        success: function (result) {
            $(".formclass").hide();
            $("#sucessmsg").show();
            document.getElementById("overlay").style.display = "none";
        },
        error: function (error) {
            console.log(error);
            document.getElementById("overlay").style.display = "none";

        }
    })
}

//send respond
function responduser() {

    var path = localStorage.getItem("DocUrl");;
    var emailids = [];
    $(".select2-selection__choice__display").each(function (data) {
        emailids.push($(this).text().split("(")[1].split(")")[0]);
    });
    var msg = document.getElementById("txtmessage").value;

    var token = localStorage.getItem("BTServiceToken");

    let payLoad = JSON.stringify(
        {
            "itemPath": localStorage.getItem("ResDocUrl").split('Documents')[1],
            "emailIDs": emailids,
            "message": msg,
            "source": "AddIns",
            "documentActivityMasterID": 1
        });

    document.getElementById("overlay").style.display = "flex";
    $.ajax({
        url: "https://btserviceapiappservice.azurewebsites.net/api/Document/sendForReviewAddIns",
        headers: {
            'Authorization': 'Bearer ' + token,
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        },
        type: 'POST',
        data: payLoad,
        dataType: 'json',
        success: function (result) {         
            
            console.log(result);
            $(".formclass").hide();
            $("#sucessmsg").show();
            document.getElementById("overlay").style.display = "none";

        },
        error: function (error) {
            console.log(error);
            document.getElementById("overlay").style.display = "none";
        }
    })
    
}



function mail() {
    console.log("before");
    var emailids = document.getElementById("txtTo").value;
    var message = document.getElementById("txtMessage").value;
    var url = "https://btserviceapiappservice.azurewebsites.net/api/Configuration/sendforreview";

    var xhr = new XMLHttpRequest();
    xhr.open("POST", url);

    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("Content-Type", "application/json");
    console.log("middle");
    xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
            console.log(xhr.status);
            console.log(xhr.responseText);
        }
    };
    let referenceid = CreateGuid();

    //let emailids = document.getElementById("txtTo").value;
    //let message = document.getElementById("txtMessage").value;
    let link = "https://blueed-my.sharepoint.com/:w:/g/personal/raghunath_mn_blueed_onmicrosoft_com/EYw8Z4FuHetEqom_atHoFZIBkYcjw5Eg0Hl5QJgs2TITTw?e=wWhBTe";

    var data = `{"documentReferenceID":"` + referenceid + `","emailIDs":"` + emailids + `"
                 ,"message": "`+ message + `" ,"status":"false","link": "` + link + `"}`;
    console.log(data);
    xhr.send(data);
    console.log("after");
}

function CreateGuid() {
    function _p8(s) {
        var p = (Math.random().toString(16) + "000000000").substr(2, 8);
        return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
    }
    return _p8() + _p8(true) + _p8(true) + _p8();
}

function goback() {
    window.history.back();
}

function redirect() {
    window.location = "http://www.w3schools.com";
}

function respond() {
    var url = "https://btserviceapiappservice.azurewebsites.net/api/Configuration/sendforreview";
    var xhr = new XMLHttpRequest();
    xhr.open("POST", url);
    let link = "https://blueed-my.sharepoint.com/:w:/g/personal/raghunath_mn_blueed_onmicrosoft_com/EYw8Z4FuHetEqom_atHoFZIBkYcjw5Eg0Hl5QJgs2TITTw?e=wWhBTe";
    var data = `{"documentReferenceID":"` + CreateGuid() + `" ,"status":"true"}`;
    console.log(data);
    xhr.send(data);
}
function getFileUrl() {
    // Get the URL of the current file.
    var fileUrl;
    Office.onReady(function (info) {

        console.log(`Office.js is now ready `);
    });
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            console.log("The file hasn't been saved yet. Save the file and try again");
            showMessage("The file hasn't been saved yet. Save the file and try again");
            //write("The file hasn't been saved yet. Save the file and try again");
        }
        else {

            // document.getElementById("fileid").value = fileUrl; 
            console.log("fileurl:1" + fileUrl);
            write(fileUrl);

        }
        console.log("fileurl:2" + fileUrl);
        return fileUrl;
    });

}
function getFileUrlforreview() {
    // Get the URL of the current file.
    var fileUrl;
    console.log("getFileUrlforreview()");
    Office.onReady(function (info) {

        console.log(`Office.js is now ready `);
    });
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            console.log("The file hasn't been saved yet. Save the file and try again");
            showMessage("The file hasn't been saved yet. Save the file and try again");
            //write("The file hasn't been saved yet. Save the file and try again");
        }
        else {

            // document.getElementById("fileid").value = fileUrl; 
            console.log("fileurl:1" + fileUrl);
            geturl(fileUrl)

        }
        console.log("fileurl:2" + fileUrl);
        return fileUrl;
    });

}
function geturl(urllink) {

}
function write(message) {
    var url = "https://btserviceapiappservice.azurewebsites.net/api/Configuration/sendforreview";

    var xhr = new XMLHttpRequest();
    xhr.open("POST", url);

    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("Content-Type", "application/json");
    console.log("middle");
    xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
            console.log(xhr.status);
            console.log(xhr.responseText);
        }
    };
    let referenceid = CreateGuid();


    //let link = "https://blueed-my.sharepoint.com/:w:/g/personal/raghunath_mn_blueed_onmicrosoft_com/EQnorFFVpY9BrVi93aenl9sBUgi4EePCSOc5Zv8qy6-PTQ?e=QHLr04";

    var link = message;
    console.log(link);
    // var data = `{"documentReferenceID":"` + referenceid + `","emailIDs":"raghunath.mn@blueed.com" + ,"message":"Sample message","status":"Approved","link": "`+link+`"}`;
    var data = `{"documentReferenceID":"` + referenceid + `","emailIDs":"raghunath.mn@blueed.onmicrosoft.com"
                 ,"message": "Please review the document" ,"status":"Approved","link":"`+ link + `"}`;
    console.log(data);
    //xhr.send(data);
}
async function post() {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
        });
    });

    await Word.run(async (context) => {

        // Create a proxy object for the document.
        var thisDocument = context.document;

        // Queue a command to load the document save state (on the saved property).
        context.load(thisDocument, 'saved');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();

        if (thisDocument.saved === false) {
            // Queue a command to save this document.
            thisDocument.save();

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            await context.sync();
            console.log('Saved the document');
        } else {
            console.log('The document has not changed since the last save.');
        }
    });

}


async function largeFileUpload(client, file) {
    try {
        let options = {
            path: "/desired upload path",
            fileName: file.name,
            rangeSize: 1024 * 1024,
        };
        const uploadTask = await MicrosoftGraph.OneDriveLargeFileUploadTask.create(client, file, options);
        const response = await uploadTask.upload();
        return response;
    } catch (err) {
        throw err;
    }
}


function getResourceGuid() {
    console.log("getResourceGuid");

    Office.context.document.getSelectedResourceAsync(
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                onError(result.error);
                console.log('Error : ' + result.error)
            }
            else {
                console.log('Resource GUID: ' + result.value)
                $('#message').html('Resource GUID: ' + result.value);
            }
        }
    );
}

function onError(error) {
    $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
}

