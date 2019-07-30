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
                        },
                        "com.k2.microsoft.teams.team.archivestatus":{
                            displayName: "Archive Status Link",
                            type: "string"
                        }
                    },
                    methods: {
                        "com.k2.microsoft.teams.team.get": {
                            displayName: "Get",
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
                                        "com.k2.microsoft.teams.team.mailnickname",
                                        "com.k2.microsoft.teams.team.weburl"
                                    ]
                        },
                        "com.k2.microsoft.teams.team.create":{
                            displayName: "Create",
                            description: "Creates a new group and adds a team to the group",
                            type: "create",
                            inputs: [   "com.k2.microsoft.teams.team.displayname",   
                                        "com.k2.microsoft.teams.team.description",
                                        "com.k2.microsoft.teams.team.mailenabled",
                                        "com.k2.microsoft.teams.team.mailnickname"
                            ],
                            requiredInputs: [ "com.k2.microsoft.teams.team.displayname" ],
                            outputs: [  "com.k2.microsoft.teams.team.id",
                                        "com.k2.microsoft.teams.team.displayname",
                                        "com.k2.microsoft.teams.team.creationdate",
                                        "com.k2.microsoft.teams.team.description",
                                        "com.k2.microsoft.teams.team.email",
                                        "com.k2.microsoft.teams.team.mailenabled",
                                        "com.k2.microsoft.teams.team.mailnickname",
                                        "com.k2.microsoft.teams.team.weburl"
                                    ]

                        },
                        "com.k2.microsoft.teams.team.add":{
                            displayName: "Add to Existing Group",
                            description: "Add a team to an already existing AAD Group",
                            type: "create",
                            outputs: [  "com.k2.microsoft.teams.team.id",
                                        "com.k2.microsoft.teams.team.displayname",
                                        "com.k2.microsoft.teams.team.creationdate",
                                        "com.k2.microsoft.teams.team.description",
                                        "com.k2.microsoft.teams.team.email",
                                        "com.k2.microsoft.teams.team.mailenabled",
                                        "com.k2.microsoft.teams.team.mailnickname",
                                        "com.k2.microsoft.teams.team.weburl"
                                    ]

                        },
                        "com.k2.microsoft.teams.team.archive":{
                            displayName: "Archive",
                            description: "Archive Team",
                            type: "execute"
                        },
                        "com.k2.microsoft.teams.team.unarchive":{
                            displayName: "Unarchive",
                            description: "Unarchive Team",
                            type: "execute"
                        },
                        "com.k2.microsoft.teams.team.checkstatus":{
                            displayName: "Check Archival Status",
                            description: "Check the stauts of an archival job.",
                            type: "read",
                            parameters: {
                                "com.k2.microsoft.teams.team.archivejoblocation": {
                                    displayName: "Archive Job Location",
                                    type: "string"
                                }
                            },
                            requiredParameters: ["com.k2.microsoft.teams.team.archivejoblocation"],
                            outputs: [  "com.k2.microsoft.teams.team.id",
                                        "com.k2.microsoft.teams.team.displayname",
                                        "com.k2.microsoft.teams.team.creationdate",
                                        "com.k2.microsoft.teams.team.description",
                                        "com.k2.microsoft.teams.team.email",
                                        "com.k2.microsoft.teams.team.mailenabled",
                                        "com.k2.microsoft.teams.team.mailnickname",
                                        "com.k2.microsoft.teams.team.weburl"
                            ]
                        } 
                        
                    }
                },
                "com.k2.microsoft.teams.channel":{
                    displayName: "Channel",
                    description: "Channel",
                    properties: {
                        "com.k2.microsoft.teams.channel.id" :{
                            displayName: "id",
                            description: "id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.displayName" :{
                            displayName: "Display Name",
                            description: "Display Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.description" :{
                            displayName: "Description",
                            description: "Description",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.email" :{
                            displayName: "Email",
                            description: "Email",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.webUrl" :{
                            displayName: "Web URL",
                            description: "Web URL",
                            type: "string"
                        }
                    },
                    methods:{
                        "com.k2.microsoft.teams.channel.get": {
                            displayName: "Get Channel",
                            type: "read",
                            parameters: {
                                "com.k2.microsoft.teams.channel.displayName": {
                                    displayName: "Display Name",
                                    type: "string"
                                }
                            },                            
                            requiredParameters: ["com.k2.microsoft.teams.channel.displayName"],
                            outputs: [  "com.k2.microsoft.teams.channel.id",
                                        "com.k2.microsoft.teams.channel.displayname",
                                        "com.k2.microsoft.teams.channel.description",
                                        "com.k2.microsoft.teams.channel.email",
                                        "com.k2.microsoft.teams.channel.webUrl"
                                    ]
                        },
                        "com.k2.microsoft.teams.channel.list": {
                            displayName: "List Channels",
                            type: "list",
                            outputs: [  "com.k2.microsoft.teams.channel.id",
                                        "com.k2.microsoft.teams.channel.displayname",
                                        "com.k2.microsoft.teams.channel.description",
                                        "com.k2.microsoft.teams.channel.email",
                                        "com.k2.microsoft.teams.channel.webUrl"
                                    ]
                        },
                        "com.k2.microsoft.teams.channel.create":{
                            displayName: "Create Channel",
                            type: "create",
                            parameters: {
                                "com.k2.microsoft.teams.channel.displayName": {
                                    displayName: "Display Name",
                                    type: "string"
                                },
                                "com.k2.microsoft.teams.channel.description": {
                                    displayName: "Description",
                                    type: "string"
                                }
                            },                            
                            requiredParameters: ["com.k2.microsoft.teams.channel.displayName","com.k2.microsoft.teams.channel.description"],
                            outputs: [  "com.k2.microsoft.teams.channel.id",
                                        "com.k2.microsoft.teams.channel.displayname",
                                        "com.k2.microsoft.teams.channel.description",
                                        "com.k2.microsoft.teams.channel.email",
                                        "com.k2.microsoft.teams.channel.webUrl"
                                    ]

                        },
                        "com.k2.microsoft.teams.channel.delete":{
                            displayName: "Delete Channel",
                            type: "delete",
                            parameters: {
                                "com.k2.microsoft.teams.channel.displayName": {
                                    displayName: "Display Name",
                                    type: "string"
                                }
                            },                            
                            requiredParameters: ["com.k2.microsoft.teams.channel.displayName"]
                        }
                    }
                }/*,
                 "com.k2.microsoft.teams.Tab":{
                    displayName: "Tab",
                    description: "Tab",
                    properties:{
                    
                    },
                    methods:{}
                }         */        
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
//#region Team Methods

//Need to map out methods for archive, unarchive and check archive status

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
                    NOTE: The groupType and securityEnabled values should be hard coded in the broker for this iteration until we understand if there is a use case where those aren't enabled.
                    We will have an add team to group method for existing AAD groupss. I feel that these options are outside the knowledge area of our target users. Those are AAD Admin level questions.
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

        REQUIRED FOR OPTIONAL METHODS:
            - Look up user by UPN to get User ID Guid.
                - GET /users/:userPrincipalName

        OPTIONAL 1:
            - Add Owner to Team using User ID from previous step.
                - POST /groups/:groupID/owners/$ref
                    REQUEST:
                        {
                            "@odata.id": "https://graph.microsoft.com/beta/users/{user id guid}"
                        }

        OPTIONAL 2:
            - Add the new owner as member of the team. Owenrs are not automatically members.
                - POST /groups/:teamID/members/$ref
                    REQUEST:
                        {
                            "@odata.id": "https://graph.microsoft.com/beta/users/{user id guid}"
                        }
    */
}

function onexecuteTeamAdd(parameters: SingleRecord, properties: SingleRecord){
    /* 
        Create Team to Team enable the group.
        - POST /groups/:groupID/team 
    */
}

function onexecuteTeamList(parameters: SingleRecord, properties: SingleRecord){
    /*  
        
        List Teams 
            GET /groups?$select=id,resourceProvisioningOptions
                NOTE: Fields to return are configurable with the select parameter. 
                Results will have to be filtered on resourceProvisioningOptions array 
                that contains the value of Team in v1 of the endpoint.
        
        List Teams By User
            GET /users/:id/joinedTeams

        List Teams by current user
            GET /me/joinedTeams

        Response Example:
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
        "value": [
            {
                "id": "4ab574b6-bbc6-4a47-8059-40fa698deb08",
                "displayName": "Dev Team",
                "description": "Development Research Team",
                "isArchived": false
            }
    ]

    */

}
//#endregion

//#region Channel Methods
function onexecuteChannelCreate(parameters: SingleRecord, properties: SingleRecord){
    /*
        Add a channel to a team.

        POST /teams/:groupID/channels
            REQUEST BODY:
                {
                    "displayName": "Innovation Discussion2",
                    "description": "This channel is where we debate all future innovation plans"
                }
     */
}

function onexecuteChannelList(parameters: SingleRecord, properties: SingleRecord){
    /*
        GET /teams/:groupID/channels
    */
}

function onexecuteChannelGet(parameters: SingleRecord, properties: SingleRecord){
    /* 
        GET /teams/{group id}/channels?$filter=displayName eq '<channel display name>'
    */
}

function onexecuteChannelDelete(parameters: SingleRecord, properties: SingleRecord){
    /*
        NOTE: Might want to combine this with Get Channel to enable this method to be called in one call using the Channel display name as the value to Delete by.
        DELETE /teams/:groupID/channels/:channelID
    */
}
//#endregion

//#region Tab Methods
function onexecuteTabCreate(parameters: SingleRecord, properties: SingleRecord){

    /*
        REFERENCE: https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs

        NOTE: We may want to create wrapper methods that call a single create Tab method to give a more
        intutive experience for K2 developers. For example: A Create OneNote Tab method, etc...

        APP IDS (entityId) for Tabs
            Word: com.microsoft.teamspace.tab.file.staticviewer.word
            Excel: com.microsoft.teamspace.tab.file.staticviewer.excel
            PowerPoint: com.microsoft.teamspace.tab.file.staticviewer.powerpoint
            PDF: com.microsoft.teamspace.tab.file.staticviewer.pdf
            OneNote: 0d820ecd-def2-4297-adad-78056cde7c78
            Planner: com.microsoft.teamspace.tab.planner
            Sharepoint: 2a527703-1f6f-4559-a332-d8a7d288cd88
            Microsoft Forms: 81fef3a6-72aa-4648-a763-de824aeafb7d
            Microsoft Stream: com.microsoftstream.embed.skypeteamstab
            Website: com.microsoft.teamspace.tab.web
            
            configuration:
                {
                "entityId": "string",
                "contentUrl": "string (HTTPS Url)",
                "websiteUrl": "string (HTTPS Url)",
                "removeUrl": "string (HTTPS Url)"  
                }

            The following are not configurable via graph and will have to be configured in teams after they are added.
            The configuration section will be omitted from the request body.
            
                Wiki: com.microsoft.teamspace.tab.wiki 
                PowerBI: om.microsoft.teamspace.tab.powerbi
                SharePoint page and lists: 2a527703-1f6f-4559-a332-d8a7d288cd88
             

        POST /teams/{group id}/channels/{channel id}/tabs
            REQUEST BODY:
            {  
                "id": "string",
                "displayName": "string",
                "webUrl": "string",

                "configuration" : <see above>
            }

    */
}

function onexecuteTabDelete(parameters: SingleRecord, properties: SingleRecord){
    /*
        DELETE /teams/:groupId/channels/:channelID/tabs/:tabID
    */
}

function onexecuteTabGet(parameters: SingleRecord, properties: SingleRecord){
    /* 
        GET /teams/{group id}/channels/{channel id}/tabs$filter=displayName eq '<tab display name>'
    */
}

function onexecuteTabList(parameters: SingleRecord, properties: SingleRecord){
    /* 
        GET /teams/{group id}/channels/{channel id}/tabs
    */
}
//#endregion


