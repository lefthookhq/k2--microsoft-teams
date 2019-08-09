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
        //case "com.k2.microsoft.teams.team.get": onexecuteTeamGet(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}


