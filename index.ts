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
                        "com.k2.microsoft.teams.team.displayname" :{
                            displayName: "Display Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.creationdate" :{
                            displayName: "Created On",
                            type: "dateTime"
                        },
                        "com.k2.microsoft.teams.team.description" :{
                            displayName: "Description",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.email" :{
                            displayName: "Email",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.mailenabled" :{
                            displayName: "Mail Enabled",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.mailnickname" :{
                            displayName: "Mail Nick Name",
                            type: "string"
                        }
                    },
                    methods: {
                        "com.k2.microsoft.teams.team.url.get": {
                            displayName: "Get Team Url",
                            type: "read",
                            parameters: {
                                "com.k2.microsoft.teams.team.teamid": {
                                    displayName: "Team Id",
                                    type: "string"
                                }
                            },                            
                            requiredParameters: ["com.k2.microsoft.teams.team.teamid"],
                            outputs: [ "com.k2.microsoft.teams.team.id", "com.k2.microsoft.teams.team.weburl"]
                        },
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
                            outputs: [  "com.k2.microsoft.teams.team.id",
                                        "com.k2.microsoft.teams.team.displayname",
                                        "com.k2.microsoft.teams.team.creationdate",
                                        "com.k2.microsoft.teams.team.description",
                                        "com.k2.microsoft.teams.team.email",
                                        "com.k2.microsoft.teams.team.mailenabled",
                                        "com.k2.microsoft.teams.team.mailnickname"
                                    ]
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
        case "com.k2.microsoft.teams.team.url.get": onexecuteTeamGetUrl(parameters, properties); break;
        case "com.k2.microsoft.teams.team.get": onexecuteTeamGet(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteTeamGetUrl(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);
        console.log(xhr.responseText);

        var obj = JSON.parse(xhr.responseText);
        postResult({
            "com.k2.microsoft.teams.team.id": obj.id,
            "com.k2.microsoft.teams.team.weburl": obj.webUrl,
            "com.k2.microsoft.teams.team.isarchived": obj.isArchived
        });  
    };
    xhr.withCredentials = true;
    xhr.open("GET", baseUriEndpoint  + "/teams/" + parameters["com.k2.microsoft.teams.team.teamid"]);
    xhr.send();
     
}

function onexecuteTeamGet(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);
        console.log(xhr.responseText);

        var obj = JSON.parse(xhr.responseText);
        postResult({
            "com.k2.microsoft.teams.team.id": obj.id,
            "com.k2.microsoft.teams.team.displayname" : obj.displayName,
            "com.k2.microsoft.teams.team.creationdate" : obj.createdDateTime,
            "com.k2.microsoft.teams.team.description" : obj.description,
            "com.k2.microsoft.teams.team.email" : obj.mail,
            "com.k2.microsoft.teams.team.mailenabled" : obj.mailEnabled,
            "com.k2.microsoft.teams.team.mailnickname" : obj.mailNickname
        });  
    };
    xhr.withCredentials = true;
    xhr.open("GET", baseUriEndpoint  + "/groups/" + parameters["com.k2.microsoft.teams.team.teamid"]);
    xhr.send();
     
}

function onexecuteTeamCreate(parameters: SingleRecord, properties: SingleRecord){
    /*
        This method should combine two Graph API calls to simplify the experence for the K2 developer. You will need the ID in the return for the next call.
            - Create group to create the group for the team. 
                - POST /groups
                    REQUEST:
                    {
                        "description": "We are the new Team",
                        "displayName": "New Team",
                        "groupTypes": [
                            "Unified"
                        ],
                        "mailEnabled": true,
                        "mailNickname": "newteam",
                        "securityEnabled": false
                    }
            - Create Team to Team enable the group.
                - POST /groups/:groupID/team
                    NOTE: No Request Body is required. Just call the post to this endpoint with the group id from the first call.

        NOTE: The user identity adding the group and team is added as the owner to the team. Could need to add a user as a team owner if additional processing isn't done as the identity that created the group.
        This next method should be optional to add a owner to a team and add them as a member during team creations.
            - Look up user by UPN to get User ID Guid.
                - GET /users/:userIDorPrincipalName
            - Add Owner to Team using User ID from previous step.
                - POST /groups/:groupID/owners/$ref
                    REQUEST:
                        {
                            "@odata.id": "https://graph.microsoft.com/beta/users/{user id guid}"
                        }
            - Optional: Add the new owner as member of the team. Owenrs are not automatically members.
                - POST /groups/:teamID/members/$ref
                    REQUEST:
                        {
                            "@odata.id": "https://graph.microsoft.com/beta/users/{user id guid}"
                        }
    */
}