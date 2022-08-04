// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
function test() {
    window.open("ShareDocument", "hello", `toolbar=no,directories=no,titlebar=no,scrollbars=no,resizable=no,status=no,location=no,toolbar=no,menubar=no,
width=500,height=500,left=300,top=300`);
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
    var data = `{"documentReferenceID":"` + referenceid + `","emailIDs":"` + emailids +`"
                 ,"message": "`+ message +`" ,"status":"false","link": "`+link +`"}`;
    console.log(data);
   // xhr.send(data);
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