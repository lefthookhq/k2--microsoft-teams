// Constants used in place of service keys.
const baseUriEndpoint = "https://graph.microsoft.com/v1.0";

ondescribe = function() {
    postSchema({
        "com.k2.microsoft.teams": {
            displayName: "Microsoft Teams",
            description: "A service for integrating with Microsoft Teams.",
            objects: {
                "com.k2.microsoft.teams.team": {
                    displayName: "Team",
                    description: "Team",
                    properties: {
                        "com.k2.microsoft.teams.team.id" :{
                            displayName: "Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.weburl" :{
                            displayName: "Web Url",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.isarchived" :{
                            displayName: "Is Archived",
                            type: "boolean"
                        }
                    },
                    methods: {
                        "com.k2.microsoft.teams.team.get": {
                            displayName: "Get Team",
                            type: "read",
                            parameters: {
                                "com.k2.microsoft.teams.team.teamid": {
                                    displayName: "Team Id",
                                    type: "string"
                                }
                            },                            
                            requiredParameters: ["com.k2.microsoft.teams.team.teamid"],
                            outputs: [ "com.k2.microsoft.teams.team.id", "com.k2.microsoft.teams.team.weburl","com.k2.microsoft.teams.team.isarchived"]
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
        case "com.k2.microsoft.teams.team": onexecuteMethods(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteMethods(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName)
    {
        case "com.k2.microsoft.teams.team.get": onexecuteTeamGet(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteTeamGet(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);
        console.log(xhr.responseText);

        var obj = JSON.parse(xhr.responseText);
        postResult({
            "com.k2.microsoft.teams.team.id": obj.Id,
            "com.k2.microsoft.teams.team.weburl": obj.webUrl,
            "com.k2.microsoft.teams.team.name": obj.isArchived
        });  
    };
    xhr.withCredentials = true;
    xhr.open("GET", baseUriEndpoint  + "/team/" + parameters["com.k2.microsoft.teams.team.teamid"]);
    xhr.send();
     
}