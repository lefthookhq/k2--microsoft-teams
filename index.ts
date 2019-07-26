// Constants used in place of service keys.
const baseUriEndpoint = "https://api.echosign.com/api/rest/v6";

ondescribe = function() {
    postSchema({
        "com.k2.microsoft.teams": {
            displayName: "Microsoft Teams",
            description: "A service for integrating with Microsoft Teams.",
            objects: {
                "com.k2.microsoft.teams.uris": {
                    displayName: "Uris",
                    description: "Base Uris for Adobe Sign API",
                    properties: {
                        "com.k2.microsoft.teams.uris.apiaccesspoint" :{
                            displayName: "API Acces Point",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.uris.webaccesspoint" :{
                            displayName: "Web Acces Point",
                            type: "string"
                        }
                    },
                    methods: {
                        "com.k2.microsoft.teams.uris.get": {
                            displayName: "Get Base Uris",
                            type: "read",
                            outputs: [ "com.k2.microsoft.teams.uris.apiaccesspoint", "com.k2.microsoft.teams.uris.webaccesspoint"]
                        }
                    }
                }                
            }
        }
    });
}

onexecute = function(objectName, methodName, parameters, properties) {
    switch (objectName)
    {
        case "com.k2.microsoft.teams.uris": onexecuteMethods(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteMethods(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName)
    {
        case "com.k2.microsoft.teams.uris.get": onexecuteUrisGet(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}





function onexecuteUrisGet(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);
        console.log(xhr.responseText);

        var obj = JSON.parse(xhr.responseText);
        postResult({
            "com.k2.microsoft.teams.uris.apiaccesspoint": obj.apiAccessPoint,
            "com.k2.microsoft.teams.uris.webaccesspoint": obj.webAccessPoint
        });  
    };
    xhr.withCredentials = true;
    xhr.open("GET", baseUriEndpoint  + "/baseUris");
    xhr.send();
     
}

//Adobe Sign Specific Methods
function getBaseUri(){
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);
        console.log(xhr.responseText);

        var obj = JSON.parse(xhr.responseText);
        apiAccessPoint: obj.apiAccessPoint;
        webAccessPoint: obj.webAccessPoint;
    };
    xhr.withCredentials = true;
    xhr.open("GET", baseUriEndpoint  + "/baseUris");
    xhr.send();
}