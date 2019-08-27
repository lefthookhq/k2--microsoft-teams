// Constants used in place of service keys.
const baseUriEndpoint = "https://graph.microsoft.com/v1.0";

ondescribe = function () {
    postSchema({
        "com.k2.microsoft.teams": {
            displayName: "Microsoft Teams",
            description: "A service for integrating with Microsoft Teams.",
            objects: {
                "com.k2.microsoft.teams.team": {
                    displayName: "Team",
                    description: "Team",
                    properties: {
                        "com.k2.microsoft.teams.team.id": {
                            displayName: "Team Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.weburl": {
                            displayName: "Web Url",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.displayname": {
                            displayName: "Display Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.creationdate": {
                            displayName: "Created On",
                            type: "dateTime"
                        },
                        "com.k2.microsoft.teams.team.description": {
                            displayName: "Description",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.email": {
                            displayName: "Email",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.mailenabled": {
                            displayName: "Mail Enabled",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.mailnickname": {
                            displayName: "Mail Nick Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.archivestatus": {
                            displayName: "Archive Status",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.operationid": {
                            displayName: "Operation Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.operationtype": {
                            displayName: "Operation Type",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.lastactiondate": {
                            displayName: "Last Action Date",
                            type: "dateTime"
                        },
                        "com.k2.microsoft.teams.team.attemptscount": {
                            displayName: "Attempts Count",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.targetresourceid": {
                            displayName: "Target Resource Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.targetresourcelocation": {
                            displayName: "Target Resource Location",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.archiveerror": {
                            displayName: "Error",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.userprincipalname": {
                            displayName: "User Principal Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.resourceprovisioningoptions": {
                            displayName: "Resource Provisioning Options",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.team.isarchived": {
                            displayName: "Is Archived",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.issuccessful": {
                            displayName: "Is Successful",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.requeststatusurl": {
                            displayName: "Request Status Url",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msallowcreateupdatechannels": {
                            displayName: "Allow create/update channels to members",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msallowdeletechannels": {
                            displayName: "Allow delete channels to members",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msallowaddremoveapps": {
                            displayName: "Allow add/remove apps to members",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msallowcreateupdateremovetabs": {
                            displayName: "Allow create/update/remove tabs to members",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msallowcreateupdateremoveconnectors": {
                            displayName: "Allow create/update/remove connectors to members",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.gsallowcreateupdatechannels": {
                            displayName: "Allow create/update channels to guests",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.gsallowdeletechannels": {
                            displayName: "Allow delete channels to guests",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msgsallowusereditmessages": {
                            displayName: "Allow user to edit messages",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msgsallowuserdeletemessages": {
                            displayName: "Allow user to delete messages",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msgsallowownerdeletemessages": {
                            displayName: "Allow owner to delete messages",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msgsallowteammentions": {
                            displayName: "Allow team mentions",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.msgsallowchannelmentions": {
                            displayName: "Allow channel mentions",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.fsallowgiphy": {
                            displayName: "Allow Giphy",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.fsallowstickersandmemes": {
                            displayName: "Allow stickers and memes",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.team.fsallowcustommemes": {
                            displayName: "Allow custom memes",
                            type: "boolean"
                        }
                    },
                    methods: {
                        "com.k2.microsoft.teams.team.get": {
                            displayName: "Get",
                            type: "read",
                            inputs: ["com.k2.microsoft.teams.team.id"],
                            requiredInputs: ["com.k2.microsoft.teams.team.id"],
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.creationdate",
                                "com.k2.microsoft.teams.team.description",
                                "com.k2.microsoft.teams.team.email",
                                "com.k2.microsoft.teams.team.mailenabled",
                                "com.k2.microsoft.teams.team.mailnickname",
                                "com.k2.microsoft.teams.team.weburl",
                                "com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.create": {
                            displayName: "Create",
                            description: "Creates a new group and adds a team to the group",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.description",
                                "com.k2.microsoft.teams.team.mailenabled",
                                "com.k2.microsoft.teams.team.mailnickname",
                                "com.k2.microsoft.teams.team.userprincipalname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.team.displayname"],
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.creationdate",
                                "com.k2.microsoft.teams.team.description",
                                "com.k2.microsoft.teams.team.email",
                                "com.k2.microsoft.teams.team.mailenabled",
                                "com.k2.microsoft.teams.team.mailnickname",
                                "com.k2.microsoft.teams.team.weburl",
                                "com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.add": {
                            displayName: "Add a team to Existing Group",
                            description: "Add a team to an already existing AAD Group",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.userprincipalname"
                            ],
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.creationdate",
                                "com.k2.microsoft.teams.team.description",
                                "com.k2.microsoft.teams.team.email",
                                "com.k2.microsoft.teams.team.mailenabled",
                                "com.k2.microsoft.teams.team.mailnickname",
                                "com.k2.microsoft.teams.team.weburl",
                                "com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.archive": {
                            displayName: "Archive",
                            description: "Archive Team",
                            type: "execute",
                            inputs: ["com.k2.microsoft.teams.team.id"],
                            requiredInputs: ["com.k2.microsoft.teams.team.id"],
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.requeststatusurl",
                                "com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.unarchive": {
                            displayName: "Unarchive",
                            description: "Unarchive Team",
                            type: "execute",
                            inputs: ["com.k2.microsoft.teams.team.id"],
                            requiredInputs: ["com.k2.microsoft.teams.team.id"],
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.requeststatusurl",
                                "com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.checkstatus": {
                            displayName: "Check Archival Status",
                            description: "Check the stauts of an archival job.",
                            type: "execute",
                            parameters: {
                                "com.k2.microsoft.teams.team.archiveoperationurl": {
                                    displayName: "Archive/Unarchive operation URL",
                                    type: "string"
                                }
                            },
                            requiredParameters: ["com.k2.microsoft.teams.team.archiveoperationurl"],
                            outputs: ["com.k2.microsoft.teams.team.operationid",
                                "com.k2.microsoft.teams.team.operationtype",
                                "com.k2.microsoft.teams.team.creationdate",
                                "com.k2.microsoft.teams.team.archivestatus",
                                "com.k2.microsoft.teams.team.lastactiondate",
                                "com.k2.microsoft.teams.team.attemptscount",
                                "com.k2.microsoft.teams.team.targetresourceid",
                                "com.k2.microsoft.teams.team.targetresourcelocation",
                                "com.k2.microsoft.teams.team.archiveerror"
                            ]
                        },
                        "com.k2.microsoft.teams.team.addmember": {
                            displayName: "Add Member",
                            description: "Add member to a team/group",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.userprincipalname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.userprincipalname"
                            ],
                            outputs: ["com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.update": {
                            displayName: "Update",
                            description: "Updates a team settings",
                            type: "update",
                            inputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.msallowcreateupdatechannels",
                                "com.k2.microsoft.teams.team.msallowdeletechannels",
                                "com.k2.microsoft.teams.team.msallowaddremoveapps",
                                "com.k2.microsoft.teams.team.msallowcreateupdateremovetabs",
                                "com.k2.microsoft.teams.team.msallowcreateupdateremoveconnectors",
                                "com.k2.microsoft.teams.team.gsallowcreateupdatechannels",
                                "com.k2.microsoft.teams.team.gsallowdeletechannels",
                                "com.k2.microsoft.teams.team.msgsallowusereditmessages",
                                "com.k2.microsoft.teams.team.msgsallowuserdeletemessages",
                                "com.k2.microsoft.teams.team.msgsallowownerdeletemessages",
                                "com.k2.microsoft.teams.team.msgsallowteammentions",
                                "com.k2.microsoft.teams.team.msgsallowchannelmentions",
                                "com.k2.microsoft.teams.team.fsallowgiphy",
                                "com.k2.microsoft.teams.team.fsallowstickersandmemes",
                                "com.k2.microsoft.teams.team.fsallowcustommemes"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.team.id"],
                            outputs: ["com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.clone": {
                            displayName: "Clone",
                            description: "Makes a copy of an existing Team",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.description",
                                "com.k2.microsoft.teams.team.mailnickname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.description",
                                "com.k2.microsoft.teams.team.mailnickname"
                            ],
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.requeststatusurl",
                                "com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.addowner": {
                            displayName: "Add Owner",
                            description: "Add owner to a team/group",
                            type: "execute",
                            parameters: {
                                "com.k2.microsoft.teams.team.addasmemberalso": {
                                    displayName: "Add as Member Also",
                                    type: "boolean"
                                }
                            },
                            inputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.userprincipalname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.userprincipalname"
                            ],
                            outputs: ["com.k2.microsoft.teams.team.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.team.list": {
                            displayName: "List all teams/groups",
                            description: "List all teams/groups",
                            type: "list",
                            parameters: {
                                "com.k2.microsoft.teams.team.displaynamestartswith": {
                                    displayName: "Display Name Starts With",
                                    type: "string"
                                }
                            },
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.resourceprovisioningoptions"
                            ]
                        },
                        "com.k2.microsoft.teams.team.myteamslist": {
                            displayName: "List my teams",
                            description: "List my teams",
                            type: "list",
                            outputs: ["com.k2.microsoft.teams.team.id",
                                "com.k2.microsoft.teams.team.displayname",
                                "com.k2.microsoft.teams.team.description",
                                "com.k2.microsoft.teams.team.isarchived"
                            ]
                        }
                    }
                },
                "com.k2.microsoft.teams.channel": {
                    displayName: "Channel",
                    description: "Channel",
                    properties: {
                        "com.k2.microsoft.teams.channel.id": {
                            displayName: "Channel Id",
                            description: "Channel Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.displayname": {
                            displayName: "Display Name",
                            description: "Display Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.description": {
                            displayName: "Description",
                            description: "Description",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.email": {
                            displayName: "Email",
                            description: "Email",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.weburl": {
                            displayName: "Web URL",
                            description: "Web URL",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.issuccessful": {
                            displayName: "Is Successful",
                            description: "Is Successful",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.channel.teamid": {
                            displayName: "Team Id",
                            description: "Team Id",
                            type: "string"
                        },
                    },
                    methods: {
                        "com.k2.microsoft.teams.channel.get": {
                            displayName: "Get Channel",
                            type: "read",
                            inputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.teamid"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.teamid"],
                            outputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.displayname",
                                "com.k2.microsoft.teams.channel.description",
                                "com.k2.microsoft.teams.channel.email",
                                "com.k2.microsoft.teams.channel.weburl",
                                "com.k2.microsoft.teams.channel.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.channel.list": {
                            displayName: "List Channels",
                            type: "list",
                            inputs: ["com.k2.microsoft.teams.channel.teamid"],
                            requiredInputs: ["com.k2.microsoft.teams.channel.teamid"],
                            outputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.displayname",
                                "com.k2.microsoft.teams.channel.description",
                                "com.k2.microsoft.teams.channel.email"
                            ]
                        },
                        "com.k2.microsoft.teams.channel.create": {
                            displayName: "Create Channel",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.channel.teamid",
                                "com.k2.microsoft.teams.channel.displayname",
                                "com.k2.microsoft.teams.channel.description",
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.channel.teamid",
                                "com.k2.microsoft.teams.channel.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.displayname",
                                "com.k2.microsoft.teams.channel.description",
                                "com.k2.microsoft.teams.channel.email",
                                "com.k2.microsoft.teams.channel.weburl",
                                "com.k2.microsoft.teams.channel.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.channel.delete": {
                            displayName: "Delete Channel",
                            type: "delete",
                            inputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.teamid"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.teamid"],
                            outputs: [
                                "com.k2.microsoft.teams.channel.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.channel.update": {
                            displayName: "Update Channel",
                            type: "update",
                            inputs: ["com.k2.microsoft.teams.channel.teamid",
                                "com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.displayname",
                                "com.k2.microsoft.teams.channel.description"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.channel.id",
                                "com.k2.microsoft.teams.channel.teamid",
                                "com.k2.microsoft.teams.channel.displayname",
                                "com.k2.microsoft.teams.channel.description"],
                            outputs: ["com.k2.microsoft.teams.channel.issuccessful"]
                        },
                    }
                },
                "com.k2.microsoft.teams.tab": {
                    displayName: "Tabs",
                    description: "Tabs",
                    properties: {
                        "com.k2.microsoft.teams.tab.id": {
                            displayName: "Tab Id",
                            description: "Tab Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.displayname": {
                            displayName: "Display Name",
                            description: "Display Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.config.entityid": {
                            displayName: "Entity Id",
                            description: "Entity Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.config.contenturl": {
                            displayName: "Content URL",
                            description: "Content URL",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.config.websiteurl": {
                            displayName: "Website URL",
                            description: "Website URL",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.config.removeurl": {
                            displayName: "Remove URL",
                            description: "Remove URL",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.teamsapp.appid": {
                            displayName: "App Id",
                            description: "App Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.teamsapp.externalid": {
                            displayName: "External Id",
                            description: "External Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.teamsapp.appdisplayname": {
                            displayName: "App Display Name",
                            description: "App Display Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.teamsapp.distmethod": {
                            displayName: "Distribution Method",
                            description: "Distribution Method",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.sortorderindex": {
                            displayName: "Sort Order Index",
                            description: "Sort Order Index",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.weburl": {
                            displayName: "Web URL",
                            description: "Web URL",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.issuccessful": {
                            displayName: "Is Successful",
                            description: "Is Successful",
                            type: "boolean"
                        },
                        "com.k2.microsoft.teams.tab.teamid": {
                            displayName: "Team Id",
                            description: "Team Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.tab.channelid": {
                            displayName: "Channel Id",
                            description: "Channel Id",
                            type: "string"
                        }
                    },
                    methods: {
                        "com.k2.microsoft.teams.tab.get": {
                            displayName: "Get tab",
                            type: "read",
                            inputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.teamsapp.appid",
                                "com.k2.microsoft.teams.tab.teamsapp.externalid",
                                "com.k2.microsoft.teams.tab.teamsapp.appdisplayname",
                                "com.k2.microsoft.teams.tab.teamsapp.distmethod",
                                "com.k2.microsoft.teams.tab.sortorderindex",
                                "com.k2.microsoft.teams.tab.weburl"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.list": {
                            displayName: "List tabs",
                            type: "list",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createwordtab": {
                            displayName: "Create MS Word tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createexceltab": {
                            displayName: "Create MS Excel tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createpowerpointtab": {
                            displayName: "Create MS Powerpoint tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createpdftab": {
                            displayName: "Create PDF tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createonenotetab": {
                            displayName: "Create One Note tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createplannertab": {
                            displayName: "Create Planner tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createsharepointtab": {
                            displayName: "Create SharePoint tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createmsformstab": {
                            displayName: "Create MS Forms tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createmsstreamtab": {
                            displayName: "Create MS Stream tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createwebsitetab": {
                            displayName: "Create Website tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createwikitab": {
                            displayName: "Create Wiki tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createpowerbitab": {
                            displayName: "Create PowerBI tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createdocumentlibrarytab": {
                            displayName: "Create Document Library tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.createcustomtab": {
                            displayName: "Create custom tab",
                            type: "create",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.teamsapp.appid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.teamsapp.appid",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname",
                                "com.k2.microsoft.teams.tab.weburl",
                                "com.k2.microsoft.teams.tab.config.entityid",
                                "com.k2.microsoft.teams.tab.config.contenturl",
                                "com.k2.microsoft.teams.tab.config.websiteurl",
                                "com.k2.microsoft.teams.tab.config.removeurl",
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.delete": {
                            displayName: "Delete tab",
                            type: "delete",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.id"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.id"
                            ],
                            outputs: [
                                "com.k2.microsoft.teams.tab.issuccessful"
                            ]
                        },
                        "com.k2.microsoft.teams.tab.update": {
                            displayName: "Update tab",
                            type: "update",
                            inputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.id",
                                "com.k2.microsoft.teams.tab.displayname"
                            ],
                            requiredInputs: ["com.k2.microsoft.teams.tab.teamid",
                                "com.k2.microsoft.teams.tab.channelid",
                                "com.k2.microsoft.teams.tab.id"
                            ],
                            outputs: ["com.k2.microsoft.teams.tab.issuccessful"]
                        },
                    }
                },
                "com.k2.microsoft.teams.apps": {
                    displayName: "Apps",
                    description: "Apps",
                    properties: {
                        "com.k2.microsoft.teams.apps.id": {
                            displayName: "App Id",
                            description: "App Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.apps.teamid": {
                            displayName: "Team Id",
                            description: "Team Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.apps.displayname": {
                            displayName: "App Display Name",
                            description: "App Display Name",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.apps.version": {
                            displayName: "version",
                            description: "version",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.apps.teamsappdefinitionid": {
                            displayName: "Teams Apps Definition Id",
                            description: "Teams Apps Definition Id",
                            type: "string"
                        },
                        "com.k2.microsoft.teams.apps.teamsappid": {
                            displayName: "Teams Apps Id",
                            description: "Teams Apps Id",
                            type: "string"
                        }
                    },

                    methods: {
                        "com.k2.microsoft.teams.apps.list": {
                            displayName: "List installed apps",
                            type: "list",
                            inputs: ["com.k2.microsoft.teams.apps.teamid"],
                            requiredInputs: ["com.k2.microsoft.teams.apps.teamid"],
                            outputs: ["com.k2.microsoft.teams.apps.id",
                                "com.k2.microsoft.teams.apps.displayname",
                                "com.k2.microsoft.teams.apps.version",
                                "com.k2.microsoft.teams.apps.teamsappdefinitionid",
                                "com.k2.microsoft.teams.apps.teamsappid"
                            ]
                        }
                    }
                }
            }
        }
    });
}

onexecute = function (objectName, methodName, parameters, properties) {
    switch (objectName) {
        case "com.k2.microsoft.teams.team": onexecuteTeam(methodName, parameters, properties); break;
        case "com.k2.microsoft.teams.channel": onexecuteChannel(methodName, parameters, properties); break;
        case "com.k2.microsoft.teams.tab": onexecuteTab(methodName, parameters, properties); break;
        case "com.k2.microsoft.teams.apps": onexecuteApp(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteApp(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case "com.k2.microsoft.teams.apps.list": onexecuteInstalledAppsList(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecuteTeam(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    try {
        parameters["com.k2.microsoft.teams.team.issuccessful"] = true;

        switch (methodName) {
            case "com.k2.microsoft.teams.team.get": onexecuteTeamGet(parameters, properties); break;
            case "com.k2.microsoft.teams.team.create": onexecuteTeamCreate(parameters, properties); break;
            case "com.k2.microsoft.teams.team.add": onexecuteTeamAdd(parameters, properties); break;
            case "com.k2.microsoft.teams.team.update": onexecuteTeamUpdate(parameters, properties); break;
            case "com.k2.microsoft.teams.team.list": onexecuteTeamList(parameters, properties); break;
            case "com.k2.microsoft.teams.team.archive": onexecuteTeamArchive(parameters, properties); break;
            case "com.k2.microsoft.teams.team.unarchive": onexecuteTeamUnarchive(parameters, properties); break;
            case "com.k2.microsoft.teams.team.addmember": onexecuteTeamAddMember(parameters, properties); break;
            case "com.k2.microsoft.teams.team.clone": onexecuteTeamClone(parameters, properties); break;
            case "com.k2.microsoft.teams.team.addowner": onexecuteTeamAddOwner(parameters, properties); break;
            case "com.k2.microsoft.teams.team.myteamslist": onexecuteTeamMyTeamsList(parameters, properties); break;
            case "com.k2.microsoft.teams.team.checkstatus": onexecuteTeamCheckStatus(parameters, properties); break;
            default: throw new Error("The method " + methodName + " is not supported..");
        }
    }
    catch (errMsg) {
        postResult({
            "com.k2.microsoft.teams.team.issuccessful": false
        });
    }
}

function onexecuteTab(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    try {
        parameters["com.k2.microsoft.teams.tab.issuccessful"] = true;

        switch (methodName) {
            case "com.k2.microsoft.teams.tab.get": onexecuteTabGet(parameters, properties); break;
            case "com.k2.microsoft.teams.tab.update": onexecuteTabUpdate(parameters, properties); break;
            case "com.k2.microsoft.teams.tab.list": onexecuteTabList(parameters, properties); break;
            case "com.k2.microsoft.teams.tab.delete": onexecuteTabDelete(parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createwordtab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createexceltab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createpowerpointtab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createpdftab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createonenotetab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createplannertab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createsharepointtab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createmsformstab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createmsstreamtab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createwebsitetab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createwikitab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createpowerbitab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createdocumentlibrarytab": onexecuteTabCreate(methodName, parameters, properties); break;
            case "com.k2.microsoft.teams.tab.createcustomtab": onexecuteTabCreate(methodName, parameters, properties); break;
            default: throw new Error("The method " + methodName + " is not supported..");
        }
    }
    catch (errMsg) {
        postResult({
            "com.k2.microsoft.teams.tab.issuccessful": false
        });
    }
}

function onexecuteTeamGet(parameters: SingleRecord, properties: SingleRecord) {

    try {
        parameters["com.k2.microsoft.teams.team.issuccessful"] = true;

        // Get Group Details By Group ID
        GetGroupDetailsById(parameters, properties, function (b) {

            parameters["com.k2.microsoft.teams.team.displayname"] = b.displayName;
            parameters["com.k2.microsoft.teams.team.creationdate"] = b.createdDateTime;
            parameters["com.k2.microsoft.teams.team.description"] = b.description;
            parameters["com.k2.microsoft.teams.team.email"] = b.mail;
            parameters["com.k2.microsoft.teams.team.mailenabled"] = b.mailEnabled;
            parameters["com.k2.microsoft.teams.team.mailnickname"] = b.mailNickname;

            //Get Team Details By Group ID
            GetTeamDetailsByID(parameters, properties, function (c) {

                parameters["com.k2.microsoft.teams.team.weburl"] = c.webUrl;

                // console.log("hello new eroor...................");
                // throw new Error("hello new eroor...................");

                //Post Results
                CreateAndReturnTeamObject(parameters, properties);
            });
        });
    }
    catch (errMsg) {
        postResult({
            "com.k2.microsoft.teams.team.issuccessful": false
        });
    }
}

function onexecuteTeamCreate(parameters: SingleRecord, properties: SingleRecord) {

    //Create a Group
    CreateGroup(parameters, properties, function (a) {

        properties["com.k2.microsoft.teams.team.id"] = parameters["com.k2.microsoft.teams.team.id"] = a.id,
        parameters["com.k2.microsoft.teams.team.creationdate"] = a.createdDateTime;
        parameters["com.k2.microsoft.teams.team.description"] = a.description;
        parameters["com.k2.microsoft.teams.team.displayname"] = a.displayName;
        parameters["com.k2.microsoft.teams.team.email"] = a.mail;
        parameters["com.k2.microsoft.teams.team.mailenabled"] = a.mailEnabled;
        parameters["com.k2.microsoft.teams.team.mailnickname"] = a.mailNickname;

        //Create a Team under the above Group
        CreateTeam(parameters, properties, function (b) {

            parameters["com.k2.microsoft.teams.team.weburl"] = b.webUrl;

            //Get User
            GetUser(parameters, properties, function (c) {


                parameters["com.k2.microsoft.teams.team.userid"] = c.id;

                //Add Group Owner
                AddGroupOwner(parameters, properties, function (d) {

                    //Add Members to the Group
                    AddGroupMembers(parameters, properties, function (d) {

                        //Post Results
                        CreateAndReturnTeamObject(parameters, properties);

                    });
                });
            });
        });
    });
}

function GetGroupIdByMailNickName(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/groups?$filter=mailNickName%20eq%20'" + String(properties["com.k2.microsoft.teams.team.mailnickname"]) + "'";

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function GetGroupDetailsById(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/groups/" + properties["com.k2.microsoft.teams.team.id"];

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function GetTeamDetailsByID(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.team.id"];

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function CreateGroup(parameters: SingleRecord, properties: SingleRecord, cb) {

    //Create Body for GROUP POST
    var data = JSON.stringify({
        "description": properties["com.k2.microsoft.teams.team.description"],
        "displayName": properties["com.k2.microsoft.teams.team.displayname"],
        "groupTypes": ["Unified"],
        "mailEnabled": properties["com.k2.microsoft.teams.team.mailenabled"],
        "mailNickname": properties["com.k2.microsoft.teams.team.mailnickname"],
        "securityEnabled": true
    });

    var url = baseUriEndpoint + "/groups/";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function CreateTeam(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = JSON.stringify({
        "memberSettings": {
            "allowCreateUpdateChannels": true
        },
        "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true
        },
        "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "strict"
        }
    });
    var url = baseUriEndpoint + "/groups/" + properties["com.k2.microsoft.teams.team.id"] + "/team";

    ExecuteRequest(url, data, "PUT", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function ArchiveTeam(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = JSON.stringify({
        "shouldSetSpoSiteReadOnlyForMembers": true
    });
    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.team.id"] + "/archive";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function UnarchiveTeam(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = null;
    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.team.id"] + "/unarchive";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function UpdateTeam(parameters: SingleRecord, properties: SingleRecord, cb) {

    //ToDo - define properties that has to be updated
    var data = JSON.stringify({
        "memberSettings": {
            "allowCreateUpdateChannels": properties["com.k2.microsoft.teams.team.msallowcreateupdatechannels"],
            "allowDeleteChannels": properties["com.k2.microsoft.teams.team.msallowdeletechannels"],
            "allowAddRemoveApps": properties["com.k2.microsoft.teams.team.msallowaddremoveapps"],
            "allowCreateUpdateRemoveTabs": properties["com.k2.microsoft.teams.team.msallowcreateupdateremovetabs"],
            "allowCreateUpdateRemoveConnectors": properties["com.k2.microsoft.teams.team.msallowcreateupdateremoveconnectors"]
        },
        "guestSettings": {
            "allowCreateUpdateChannels": properties["com.k2.microsoft.teams.team.gsallowcreateupdatechannels"],
            "allowDeleteChannels": properties["com.k2.microsoft.teams.team.gsallowdeletechannels"]
        },
        "messagingSettings": {
            "allowUserEditMessages": properties["com.k2.microsoft.teams.team.msgsallowusereditmessages"],
            "allowUserDeleteMessages": properties["com.k2.microsoft.teams.team.msgsallowuserdeletemessages"],
            "allowOwnerDeleteMessages": properties["com.k2.microsoft.teams.team.msgsallowownerdeletemessages"],
            "allowTeamMentions": properties["com.k2.microsoft.teams.team.msgsallowteammentions"],
            "allowChannelMentions": properties["com.k2.microsoft.teams.team.msgsallowchannelmentions"]
        },
        "funSettings": {
            "allowGiphy": properties["com.k2.microsoft.teams.team.fsallowgiphy"],
            "allowStickersAndMemes": properties["com.k2.microsoft.teams.team.fsallowstickersandmemes"],
            "allowCustomMemes": properties["com.k2.microsoft.teams.team.fsallowcustommemes"]
        }
    });
    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.team.id"];

    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function CloneTeam(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = JSON.stringify({
        "displayName": properties["com.k2.microsoft.teams.team.displayname"],
        "description": properties["com.k2.microsoft.teams.team.description"],
        "mailNickname": properties["com.k2.microsoft.teams.team.mailnickname"],
        "partsToClone": "apps,tabs,settings,channels,members",
        "visibility": "public"
    });

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.team.id"] + "/clone";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function GetUser(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/users/" + properties["com.k2.microsoft.teams.team.userprincipalname"];

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function AddGroupOwner(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/users/" + parameters["com.k2.microsoft.teams.team.userid"]
    });

    var url = baseUriEndpoint + "/groups/" + properties["com.k2.microsoft.teams.team.id"] + "/owners/$ref";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function AddGroupMembers(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/directoryObjects/" + parameters["com.k2.microsoft.teams.team.userid"]
    });

    var url = baseUriEndpoint + "/groups/" + properties["com.k2.microsoft.teams.team.id"] + "/members/$ref";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function ExecuteRequest(url: string, data: string, requestType: string, cb) {

    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;

        if (xhr.status == 201) {
            console.log(xhr.responseText);
            var obj = JSON.parse(xhr.responseText);
            if (typeof cb === 'function')
                cb(obj);
        }
        else if (xhr.status == 204) {
            console.log(xhr.responseText);
            if (typeof cb === 'function')
                cb(xhr.responseText);
        }
        else if (xhr.status == 200) {
            console.log(xhr.responseText);
            var obj = JSON.parse(xhr.responseText);
            if (typeof cb === 'function')
                cb(obj);
        }
        else if (xhr.status == 202) {
            console.log(xhr.responseText);
            if (typeof cb === 'function')
                cb(null);
        }
        else {
            //throw new Error("Failed with status " + xhr.status);
            console.log("Failed with status " + xhr.status);

            postResult({
                "com.k2.microsoft.teams.team.issuccessful": false
            });
        }
    };

    console.log(url);
    xhr.open(requestType.toUpperCase(), url);

    // Authentication Header
    xhr.withCredentials = true;
    xhr.setRequestHeader("Accept", "application/json");

    if (requestType.toUpperCase() == "PUT" || requestType.toUpperCase() == "POST" || requestType.toUpperCase() == "PATCH") {
        xhr.setRequestHeader("Content-Type", "application/json");
    }

    xhr.send(data);
}

function CreateAndReturnTeamObject(parameters: SingleRecord, properties: SingleRecord) {

    if (String(properties["com.k2.microsoft.teams.team.id"]).length > 0)
        parameters["com.k2.microsoft.teams.team.id"] = properties["com.k2.microsoft.teams.team.id"];

    postResult({
        "com.k2.microsoft.teams.team.id": parameters["com.k2.microsoft.teams.team.id"],
        "com.k2.microsoft.teams.team.displayname": parameters["com.k2.microsoft.teams.team.displayname"],
        "com.k2.microsoft.teams.team.creationdate": parameters["com.k2.microsoft.teams.team.creationdate"],
        "com.k2.microsoft.teams.team.description": parameters["com.k2.microsoft.teams.team.description"],
        "com.k2.microsoft.teams.team.email": parameters["com.k2.microsoft.teams.team.email"],
        "com.k2.microsoft.teams.team.mailenabled": parameters["com.k2.microsoft.teams.team.mailenabled"],
        "com.k2.microsoft.teams.team.mailnickname": parameters["com.k2.microsoft.teams.team.mailnickname"],
        "com.k2.microsoft.teams.team.weburl": parameters["com.k2.microsoft.teams.team.weburl"],
        "com.k2.microsoft.teams.team.issuccessful": parameters["com.k2.microsoft.teams.team.issuccessful"]
    });

}

function onexecuteTeamAdd(parameters: SingleRecord, properties: SingleRecord) {

    //To Do - Should we make a call to Get Group Details by ID to get the team object details - for returning back to K2 user

    // Add Team to a group
    CreateTeam(parameters, properties, function (b) {

        parameters["com.k2.microsoft.teams.team.weburl"] = b.webUrl;

        // Get user
        GetUser(parameters, properties, function (c) {

            parameters["com.k2.microsoft.teams.team.userid"] = c.id;

            // Add Owner
            AddGroupOwner(parameters, properties, function (d) {

                // Add Owner As Member
                AddGroupMembers(parameters, properties, function (e) {

                    //Return Team Object
                    CreateAndReturnTeamObject(parameters, properties);
                });
            });
        });
    });
}

function onexecuteTeamUpdate(parameters: SingleRecord, properties: SingleRecord) {

    UpdateTeam(parameters, properties, function (c) {

        if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
            postResult({
                "com.k2.microsoft.teams.team.issuccessful": true
            });
        }
        //CreateAndReturnTeamObject(parameters, properties);

    })
}

function onexecuteTeamMyTeamsList(parameters: SingleRecord, properties: SingleRecord) {
    GetMyTeams(parameters, properties, function (a) {

        console.log(a);

        postResult(
            a.value.map(x => {
                return {
                    "com.k2.microsoft.teams.team.id": x.id,
                    "com.k2.microsoft.teams.team.displayname": x.displayName,
                    "com.k2.microsoft.teams.team.description": x.description,
                    "com.k2.microsoft.teams.team.isarchived": x.isArchived
                }
            }));
    });
}

function onexecuteTeamList(parameters: SingleRecord, properties: SingleRecord) {
    GetTeams(parameters, properties, function (a) {

        console.log(a);

        postResult(
            a.value.map(x => {
                return {
                    "com.k2.microsoft.teams.team.id": x.id,
                    "com.k2.microsoft.teams.team.displayname": x.displayName,
                    "com.k2.microsoft.teams.team.resourceprovisioningoptions": x.resourceProvisioningOptions[0]
                }
            }));
    });
}

function GetTeams(parameters: SingleRecord, properties: SingleRecord, cb) {

    if (parameters["com.k2.microsoft.teams.team.displaynamestartswith"] == null || parameters["com.k2.microsoft.teams.team.displaynamestartswith"] == "")
        var url = baseUriEndpoint + "/groups?$select=id,displayName,resourceProvisioningOptions";
    else
        var url = baseUriEndpoint + "/groups?$filter=startswith(displayName, '" + parameters["com.k2.microsoft.teams.team.displaynamestartswith"] + "')&$select=id,displayName,resourceProvisioningOptions";

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function GetMyTeams(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/me/joinedTeams";

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function onexecuteTeamArchive(parameters: SingleRecord, properties: SingleRecord) {

    ArchiveTeam(parameters, properties, function (b) {

        // CreateAndReturnTeamObject(parameters, properties);

        postResult({
            "com.k2.microsoft.teams.team.id": properties["com.k2.microsoft.teams.team.id"],
            "com.k2.microsoft.teams.team.requeststatusurl": parameters["com.k2.microsoft.teams.team.requeststatusurl"],
            "com.k2.microsoft.teams.team.issuccessful": parameters["com.k2.microsoft.teams.team.issuccessful"]
        });

    });
}

function onexecuteTeamUnarchive(parameters: SingleRecord, properties: SingleRecord) {

    UnarchiveTeam(parameters, properties, function (b) {

        CreateAndReturnTeamObject(parameters, properties);

        postResult({
            "com.k2.microsoft.teams.team.id": properties["com.k2.microsoft.teams.team.id"],
            "com.k2.microsoft.teams.team.requeststatusurl": parameters["com.k2.microsoft.teams.team.requeststatusurl"],
            "com.k2.microsoft.teams.team.issuccessful": parameters["com.k2.microsoft.teams.team.issuccessful"]
        });

    });
}

function CheckArchivalStatus(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = null;
    var url = baseUriEndpoint + "/" + parameters["com.k2.microsoft.teams.team.archiveoperationurl"];

    ExecuteRequest(url, data, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function onexecuteTeamCheckStatus(parameters: SingleRecord, properties: SingleRecord) {

    CheckArchivalStatus(parameters, properties, function (b) {

        postResult({
            "com.k2.microsoft.teams.team.operationid": b.id,
            "com.k2.microsoft.teams.team.operationtype": b.operationType,
            "com.k2.microsoft.teams.team.creationdate": b.createdDateTime,
            "com.k2.microsoft.teams.team.archivestatus": b.status,
            "com.k2.microsoft.teams.team.lastactiondate": b.lastActionDateTime,
            "com.k2.microsoft.teams.team.attemptscount": b.attemptsCount,
            "com.k2.microsoft.teams.team.targetresourceid": b.targetResourceId,
            "com.k2.microsoft.teams.team.targetresourcelocation": b.targetResourceLocation,
            "com.k2.microsoft.teams.team.archiveerror": b.error
        });

    });
}

function onexecuteTeamClone(parameters: SingleRecord, properties: SingleRecord) {

    CloneTeam(parameters, properties, function (b) {

        //CreateAndReturnTeamObject(parameters, properties);

        postResult({
            "com.k2.microsoft.teams.team.id": properties["com.k2.microsoft.teams.team.id"],
            "com.k2.microsoft.teams.team.requeststatusurl": parameters["com.k2.microsoft.teams.team.requeststatusurl"],
            "com.k2.microsoft.teams.team.issuccessful": parameters["com.k2.microsoft.teams.team.issuccessful"]
        });

    })
}

function onexecuteChannel(methodName: string, parameters: SingleRecord, properties: SingleRecord) {

    parameters["com.k2.microsoft.teams.channel.issuccessful"] = true;

    switch (methodName) {
        case "com.k2.microsoft.teams.channel.get": onexecuteChannelGet(parameters, properties); break;
        case "com.k2.microsoft.teams.channel.list": onexecuteChannelList(parameters, properties); break;
        case "com.k2.microsoft.teams.channel.create": onexecuteChannelCreate(parameters, properties); break;
        case "com.k2.microsoft.teams.channel.delete": onexecuteChannelDelete(parameters, properties); break;
        case "com.k2.microsoft.teams.channel.update": onexecuteChannelUpdate(parameters, properties); break;
        default: throw new Error("The channel method " + methodName + " is not supported...");
    }
}

function onexecuteTeamAddMember(parameters: SingleRecord, properties: SingleRecord) {

    GetUser(parameters, properties, function (b) {

        parameters["com.k2.microsoft.teams.team.userprincipalname"] = b.userPrincipalName;
        parameters["com.k2.microsoft.teams.team.userid"] = b.id;

        AddGroupMembers(parameters, properties, function (c) {

            //ToDO - remove the if condition and handle in try catch block
            if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                postResult({
                    "com.k2.microsoft.teams.team.issuccessful": true
                });
            }
        });
    });
}

function onexecuteTeamAddOwner(parameters: SingleRecord, properties: SingleRecord) {

    GetUser(parameters, properties, function (b) {

        parameters["com.k2.microsoft.teams.team.userprincipalname"] = b.userPrincipalName;
        parameters["com.k2.microsoft.teams.team.userid"] = b.id;

        AddGroupOwner(parameters, properties, function (c) {

            var stringValue: String = String(parameters["com.k2.microsoft.teams.team.addasmemberalso"]);
            var boolValue = stringValue.toLowerCase() == 'true' ? true : false;

            if (boolValue) {
                AddGroupMembers(parameters, properties, function (d) {

                    if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                        postResult({
                            "com.k2.microsoft.teams.team.issuccessful": true
                        });
                    }
                });
            }
            else {
                if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                    postResult({
                        "com.k2.microsoft.teams.team.issuccessful": true
                    });
                }
            }
        });
    });
}

function onexecuteChannelGet(parameters: SingleRecord, properties: SingleRecord) {

    GetChannel(parameters, properties, function (a) {

        parameters["com.k2.microsoft.teams.channel.id"] = a.id;
        parameters["com.k2.microsoft.teams.channel.displayname"] = a.displayName;
        parameters["com.k2.microsoft.teams.channel.description"] = a.description;
        parameters["com.k2.microsoft.teams.channel.email"] = a.email;
        parameters["com.k2.microsoft.teams.channel.weburl"] = a.webUrl;

        CreateAndReturnChannelObject(parameters, properties);

    });
}

function onexecuteChannelList(parameters: SingleRecord, properties: SingleRecord) {

    GetChannelList(parameters, properties, function (a) {

        postResult(
            a.value.map(x => {
                return {
                    "com.k2.microsoft.teams.channel.id": x.id,
                    "com.k2.microsoft.teams.channel.displayname": x.displayName,
                    "com.k2.microsoft.teams.channel.description": x.description,
                    "com.k2.microsoft.teams.channel.email": x.email
                }
            }));
    });
}

function onexecuteChannelCreate(parameters: SingleRecord, properties: SingleRecord) {

    CreateChannel(parameters, properties, function (a) {

        parameters["com.k2.microsoft.teams.channel.id"] = a.id;
        parameters["com.k2.microsoft.teams.channel.displayname"] = a.displayName;
        parameters["com.k2.microsoft.teams.channel.description"] = a.description;
        parameters["com.k2.microsoft.teams.channel.email"] = a.email;
        parameters["com.k2.microsoft.teams.channel.weburl"] = a.webUrl;

        CreateAndReturnChannelObject(parameters, properties);

    });
}

function onexecuteChannelUpdate(parameters: SingleRecord, properties: SingleRecord) {

    UpdateChannel(parameters, properties, function (a) {

        parameters["com.k2.microsoft.teams.channel.id"] = a.id;
        parameters["com.k2.microsoft.teams.channel.displayname"] = a.displayName;
        parameters["com.k2.microsoft.teams.channel.description"] = a.description;
        parameters["com.k2.microsoft.teams.channel.email"] = a.email;
        parameters["com.k2.microsoft.teams.channel.weburl"] = a.webUrl;

        // CreateAndReturnChannelObject(parameters, properties);

        postResult({
            "com.k2.microsoft.teams.channel.issuccessful": true
        });

    });
}

function onexecuteChannelDelete(parameters: SingleRecord, properties: SingleRecord) {

    DeleteChannel(parameters, properties, function (a) {

        postResult({
            "com.k2.microsoft.teams.channel.issuccessful": true
        });
    });
}

function DeleteChannel(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.channel.teamid"] + "/channels/" + properties["com.k2.microsoft.teams.channel.id"];

    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function CreateChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties["com.k2.microsoft.teams.channel.displayname"],
        "description": properties["com.k2.microsoft.teams.channel.description"],
        //"email":properties["com.k2.microsoft.teams.channel.email"]
    });

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.channel.teamid"] + "/channels";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function CreateAndReturnChannelObject(parameters: SingleRecord, properties: SingleRecord) {

    var ChannelId = String(parameters["com.k2.microsoft.teams.channel.id"]);
    if (ChannelId == null || ChannelId == "undefined" || ChannelId == "" || ChannelId == undefined)
        parameters["com.k2.microsoft.teams.channel.id"] = properties["com.k2.microsoft.teams.channel.id"];

    postResult({
        "com.k2.microsoft.teams.channel.id": parameters["com.k2.microsoft.teams.channel.id"],
        "com.k2.microsoft.teams.channel.displayname": parameters["com.k2.microsoft.teams.channel.displayname"],
        "com.k2.microsoft.teams.channel.description": parameters["com.k2.microsoft.teams.channel.description"],
        "com.k2.microsoft.teams.channel.email": parameters["com.k2.microsoft.teams.channel.email"],
        "com.k2.microsoft.teams.channel.weburl": parameters["com.k2.microsoft.teams.channel.weburl"],
        "com.k2.microsoft.teams.channel.issuccessful": parameters["com.k2.microsoft.teams.channel.issuccessful"]
    });

}

function GetChannel(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.channel.teamid"] + "/channels/" + properties["com.k2.microsoft.teams.channel.id"];

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function GetChannelList(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.channel.teamid"] + "/channels?$select=id, displayname, description, email";

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function UpdateChannel(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = JSON.stringify({
        "displayName": properties["com.k2.microsoft.teams.channel.displayname"],
        "description": properties["com.k2.microsoft.teams.channel.description"],
        "email": "test@k2rocks.com"
    });

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.channel.teamid"] + "/channels/" + properties["com.k2.microsoft.teams.channel.id"];

    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });

}
function onexecuteTabGet(parameters: SingleRecord, properties: SingleRecord) {

    GetTabInformation(parameters, properties, function (a) {

        postResult({
            "com.k2.microsoft.teams.tab.id": a.id,
            "com.k2.microsoft.teams.tab.displayname": a.displayName,
            "com.k2.microsoft.teams.tab.config.entityid": a.configuration.entityId,
            "com.k2.microsoft.teams.tab.config.contenturl": a.configuration.contentUrl,
            "com.k2.microsoft.teams.tab.config.websiteurl": a.configuration.websiteUrl,
            "com.k2.microsoft.teams.tab.config.removeurl": a.configuration.removeUrl,
            "com.k2.microsoft.teams.tab.teamsapp.appid": a.teamsApp.id,
            "com.k2.microsoft.teams.tab.teamsapp.externalid": a.teamsApp.externalId,
            "com.k2.microsoft.teams.tab.teamsapp.appdisplayname": a.teamsApp.displayName,
            "com.k2.microsoft.teams.tab.teamsapp.distmethod": a.teamsApp.distributionMethod,
            "com.k2.microsoft.teams.tab.sortorderindex": a.sortOrderIndex,
            "com.k2.microsoft.teams.tab.weburl": a.webUrl
        });
    });
}

function GetTabInformation(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.tab.teamid"] + "/Channels/" + properties["com.k2.microsoft.teams.tab.channelid"] + "/tabs/" + properties["com.k2.microsoft.teams.tab.id"] + "?$expand=teamsApp"

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });

}

function onexecuteTabUpdate(parameters: SingleRecord, properties: SingleRecord) {

    UpdateTab(parameters, properties, function (a) {

        postResult({
            "com.k2.microsoft.teams.tab.issuccessful": true
        });

    });

}

function UpdateTab(parameters: SingleRecord, properties: SingleRecord, cb) {

    var data = JSON.stringify({
        "displayName": properties["com.k2.microsoft.teams.tab.displayname"]
    });

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.tab.teamid"] + "/Channels/" + properties["com.k2.microsoft.teams.tab.channelid"] + "/tabs/" + properties["com.k2.microsoft.teams.tab.id"]

    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });

}

function onexecuteTabList(parameters: SingleRecord, properties: SingleRecord) {

    GetTabList(parameters, properties, function (a) {

        postResult(
            a.value.map(x => {
                return {
                    "com.k2.microsoft.teams.tab.id": x.id,
                    "com.k2.microsoft.teams.tab.displayname": x.displayName,
                    "com.k2.microsoft.teams.tab.weburl": x.webUrl
                }
            }));
    });
}

function onexecuteTabCreate(methodName: string, parameters: SingleRecord, properties: SingleRecord) {

    switch (methodName) {
        case "com.k2.microsoft.teams.tab.createwordtab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Word", properties));
            break;
        case "com.k2.microsoft.teams.tab.createexceltab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Excel", properties));
            break;
        case "com.k2.microsoft.teams.tab.createpowerpointtab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Powerpoint", properties));
            break;
        case "com.k2.microsoft.teams.tab.createpdftab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("PDF", properties));
            break;
        case "com.k2.microsoft.teams.tab.createonenotetab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("OneNote", properties));
            break;
        case "com.k2.microsoft.teams.tab.createplannertab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Planner", properties));
            break;
        case "com.k2.microsoft.teams.tab.createsharepointtab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("SharePoint", properties));
            break;
        case "com.k2.microsoft.teams.tab.createmsformstab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("MicrosoftForms", properties));
            break;
        case "com.k2.microsoft.teams.tab.createmsstreamtab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("MicrosoftStream", properties));
            break;
        case "com.k2.microsoft.teams.tab.createwebsitetab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Website", properties));
            break;
        case "com.k2.microsoft.teams.tab.createwikitab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Wiki", properties));
            break;
        case "com.k2.microsoft.teams.tab.createpowerbitab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("PowerBI", properties));
            break;
        case "com.k2.microsoft.teams.tab.createdocumentlibrarytab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("DocumentLibrary", properties));
            break;
        case "com.k2.microsoft.teams.tab.createcustomtab":
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Custom", properties));
            break;

        default: throw new Error("The object " + methodName + " is not supported.");
    }
}


function prepareDataAndCreateTab(parameters: SingleRecord, properties: SingleRecord, requestBody: string) {

    CreateTab(parameters, properties, requestBody, function (a) {
        // CreateAndReturnChannelObject(parameters, properties);

        postResult({
            "com.k2.microsoft.teams.tab.id": a.id,
            "com.k2.microsoft.teams.tab.displayname": a.displayName,
            "com.k2.microsoft.teams.tab.weburl": a.webUrl,
            "com.k2.microsoft.teams.tab.config.entityid": a.configuration.entityId,
            "com.k2.microsoft.teams.tab.config.contenturl": a.configuration.contentUrl,
            "com.k2.microsoft.teams.tab.config.websiteurl": a.configuration.websiteUrl,
            "com.k2.microsoft.teams.tab.config.removeurl": a.configuration.removeUrl,
            "com.k2.microsoft.teams.tab.issuccessful": true
        });
    });
}

function CreateTab(parameters: SingleRecord, properties: SingleRecord, data: string, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.tab.teamid"] + "/channels/" + properties["com.k2.microsoft.teams.tab.channelid"] + "/tabs";

    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}

function getRequestBody(tabType: string, properties) {

    var data;

    switch (tabType) {

        case "Word":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.word"
            });
            break;
        case "Excel":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.excel"
            });
            break;
        case "Powerpoint":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.powerpoint"
            });
            break;
        case "PDF":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.pdf"
            });
            break;
        case "OneNote":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78"
            });
            break;
        case "Planner":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
            });
            break;
        case "SharePoint":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88"
            });
            break;
        case "MicrosoftForms":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/81fef3a6-72aa-4648-a763-de824aeafb7d"
            });
            break;
        case "MicrosoftStream":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoftstream.embed.skypeteamstab"
            });
            break;
        case "Website":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"
            });
            break;
        case "Wiki":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.wiki"
            });
            break;
        case "PowerBI":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.powerbi"
            });
            break;
        case "DocumentLibrary":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.files.sharepoint"
            });
            break;
        case "Custom":
            data = JSON.stringify({
                "displayName": properties["com.k2.microsoft.teams.tab.displayname"],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + properties["com.k2.microsoft.teams.tab.teamsapp.appid"]
            });
            break;
        default: throw new Error("Tab Type is not supported or app is not installed!");
    }
    return data;
}

function onexecuteTabDelete(parameters: SingleRecord, properties: SingleRecord) {
    DeleteTab(parameters, properties, function (a) {

        if (a == null || a == "") {
            postResult({
                "com.k2.microsoft.teams.tab.issuccessful": true
            });
        }
    });
}

function DeleteTab(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.tab.teamid"] + "/Channels/" + properties["com.k2.microsoft.teams.tab.channelid"] + "/tabs/" + properties["com.k2.microsoft.teams.tab.id"]

    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });

}


function GetInstalledAppsList(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.apps.teamid"] + "/installedApps?$expand=teamsAppDefinition";

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}


function onexecuteInstalledAppsList(parameters: SingleRecord, properties: SingleRecord) {
    GetInstalledAppsList(parameters, properties, function (a) {

        postResult(
            a.value.map(x => {
                return {
                    "com.k2.microsoft.teams.apps.id": x.id,
                    "com.k2.microsoft.teams.apps.displayname": x.teamsAppDefinition.displayName,
                    "com.k2.microsoft.teams.apps.version": x.teamsAppDefinition.version,
                    "com.k2.microsoft.teams.apps.teamsappdefinitionid": x.teamsAppDefinition.id,
                    "com.k2.microsoft.teams.apps.teamsappid": x.teamsAppDefinition.teamsAppId
                }
            }));
    });

}

function GetTabList(parameters: SingleRecord, properties: SingleRecord, cb) {

    var url = baseUriEndpoint + "/teams/" + properties["com.k2.microsoft.teams.tab.teamid"] + "/channels/" + properties["com.k2.microsoft.teams.tab.channelid"] + "/tabs?$select=id,displayName,webUrl";

    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function') cb(responseText);
    });
}
