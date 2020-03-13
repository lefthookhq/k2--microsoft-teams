import '@k2oss/k2-broker-core';

metadata = {
    systemName: "MSTeamsJSServiceProvider",
    displayName: "MSTeamsJSServiceProvider",
    description: "An example broker"
};

// Constants used in place of service keys.
const baseUriEndpoint = "https://graph.microsoft.com/v1.0";
const baseUriEndpointBeta = "https://graph.microsoft.com/beta";

//
// Teams
Teams = "teams"; // "com.k2.microsoft.teams";

// "objects" off of "com.k2.microsoft.teams"
Team = "team"; //  "com.k2.microsoft.teams.team";
Channel = "channel"; //  "com.k2.microsoft.teams.channel";
Tab = "tab"; //  "com.k2.microsoft.teams.tab";
Apps = "apps"; //  "com.k2.microsoft.teams.apps";

//
// Team
TeamId = "id"; //  "com.k2.microsoft.teams.team.id";
TeamWeburl = "weburl"; //  "com.k2.microsoft.teams.team.weburl";
TeamDisplayName = "displayName"; //  "com.k2.microsoft.teams.team.displayname";
TeamCreationDate = "creationDate"; //  "com.k2.microsoft.teams.team.creationdate";
TeamDescription = "description"; //  "com.k2.microsoft.teams.team.description";
TeamEmail = "email"; //  "com.k2.microsoft.teams.team.email";
TeamMailEnabled = "mailEnabled"; //  "com.k2.microsoft.teams.team.mailenabled";
TeamMailNickname = "mailNickname"; //  "com.k2.microsoft.teams.team.mailnickname";
TeamArchiveStatus = "archiveStatus"; //  "com.k2.microsoft.teams.team.archivestatus";
TeamOperationId = "operationId"; //  "com.k2.microsoft.teams.team.operationid";
TeamOperationType = "operationType"; //  "com.k2.microsoft.teams.team.operationtype";
TeamLastActionDate = "lastActionDate"; //  "com.k2.microsoft.teams.team.lastactiondate";
TeamAttemptsCount = "attemptsCount"; //  "com.k2.microsoft.teams.team.attemptscount";
TeamTargetResourceId = "targetResourceId"; //  "com.k2.microsoft.teams.team.targetresourceid";
TeamTargetResourceLocation = "targetResourceLocation"; //  "com.k2.microsoft.teams.team.targetresourcelocation";
TeamArchiveError = "archiveError"; //  "com.k2.microsoft.teams.team.archiveerror";
TeamUserPrincipalName = "userPrincipalName"; //  "com.k2.microsoft.teams.team.userprincipalname";
TeamResourceProvisioningOptions = "resourceProvisioningOptions"; //  "com.k2.microsoft.teams.team.resourceprovisioningoptions";
TeamIsArchived = "isArchived"; //  "com.k2.microsoft.teams.team.isarchived";
TeamIsSuccessful = "isSuccessful"; //  "com.k2.microsoft.teams.team.issuccessful"
TeamRequestStatusUrl = "requestStatusUrl"; //  "com.k2.microsoft.teams.team.requeststatusurl";
TeamMsAllowCreateUpdateChannels = "msAllowCreateUpdateChannels"; //  "com.k2.microsoft.teams.team.msallowcreateupdatechannels";
TeamMsAllowDeleteChannels = "msAllowDeleteChannels"; //  "com.k2.microsoft.teams.team.msallowdeletechannels";
TeamMsAllowAddRemoveApps = "msAllowAddRemoveApps"; //  "com.k2.microsoft.teams.team.msallowaddremoveapps";
TeamMsAllowCreateUpdateRemoveTabs = "msAllowCreateUpdateRemoveTabs"; //  "com.k2.microsoft.teams.team.msallowcreateupdateremovetabs";
TeamMsAllowCreateUpdateRemoveConnectors = "msAllowCreateUpdateRemoveConnectors"; //  "com.k2.microsoft.teams.team.msallowcreateupdateremoveconnectors";
TeamGsAllowCreateUpdateChannels = "gsAllowCreateUpdateChannels"; //  "com.k2.microsoft.teams.team.gsallowcreateupdatechannels";
TeamGsAllowDeleteChannels = "gsAllowDeleteChannels"; //  "com.k2.microsoft.teams.team.gsallowdeletechannels";
TeamMsgAllowUserEditMessages = "msgsAllowUserEditMessages"; //  "com.k2.microsoft.teams.team.msgsallowusereditmessages";
TeamMsgAllowUserDeleteMessages = "msgsAllowUserDeleteMessages"; //  "com.k2.microsoft.teams.team.msgsallowuserdeletemessages";
TeamMsgAllowOwnerDeleteMessages = "msgsAllowOwnerDeleteMessages"; //  "com.k2.microsoft.teams.team.msgsallowownerdeletemessages";
TeamMsgAllowTeamMentions = "msgsAllowTeamMentions"; //  "com.k2.microsoft.teams.team.msgsallowteammentions";
TeamMsgAllowChannelMentions = "msgsAllowChannelMentions"; //  "com.k2.microsoft.teams.team.msgsallowchannelmentions";
TeamFsAllowGiphy = "fsAllowGiphy"; //  "com.k2.microsoft.teams.team.fsallowgiphy";
TeamFsAllowStickersAndMemes = "fsAllowStickersAndMemes"; //  "com.k2.microsoft.teams.team.fsallowstickersandmemes";
TeamFsAllowCustomMemes = "fsAllowCustomMemes"; //  "com.k2.microsoft.teams.team.fsallowcustommemes";

TeamGet = "get"; //  "com.k2.microsoft.teams.team.get";
TeamCreate = "create"; //  "com.k2.microsoft.teams.team.create";
TeamAdd = "add"; //  "com.k2.microsoft.teams.team.add";
TeamArchive = "archive"; //  "com.k2.microsoft.teams.team.archive";
TeamUnarchive = "unarchive"; //  "com.k2.microsoft.teams.team.unarchive";
TeamCheckStatus = "checkStatus"; //  "com.k2.microsoft.teams.team.checkstatus";
TeamAddMember = "addMember"; //  "com.k2.microsoft.teams.team.addmember";
TeamUpdate = "update"; //  "com.k2.microsoft.teams.team.update";
TeamClone = "clone"; //  "com.k2.microsoft.teams.team.clone";
TeamAddOwner = "addowner"; //  "com.k2.microsoft.teams.team.addowner";
TeamList = "list"; //  "com.k2.microsoft.teams.team.list";
TeamMyTeamsList = "myTeamsList"; //  "com.k2.microsoft.teams.team.myteamslist";

TeamArchiveOperationUrl = "archiveOperationUrl"; //  "com.k2.microsoft.teams.team.archiveoperationurl";
TeamAddAsMemberAlso = "addAsMemberAlso"; //  "com.k2.microsoft.teams.team.addasmemberalso";
TeamDisplayNameStartsWith = "displayNameStartsWith"; //  "com.k2.microsoft.teams.team.displaynamestartswith";
TeamUserId = "userId"; //  "com.k2.microsoft.teams.team.userid";

//
// Channel
ChannelId = "id"; //  "com.k2.microsoft.teams.channel.id";
ChannelDisplayName = "displayName"; //  "com.k2.microsoft.teams.channel.displayname";
ChannelDescription = "description"; //  "com.k2.microsoft.teams.channel.description";
ChannelEmail = "email"; //  "com.k2.microsoft.teams.channel.email";
ChannelWeburl = "weburl"; //  "com.k2.microsoft.teams.channel.weburl";
ChannelIsSuccessful = "isSuccessful"; //  "com.k2.microsoft.teams.channel.issuccessful";
ChannelTeamId = "teamId"; //  "com.k2.microsoft.teams.channel.teamid";
ChannelMessageSubject = "messageSubject"; //  "com.k2.microsoft.teams.channel.message.subject";
ChannelMessageBody = "messageBody"; //  "com.k2.microsoft.teams.channel.message.body";
ChannelMessageIsImportant = "messageIsImportant"; //  "com.k2.microsoft.teams.channel.message.isimportant";

ChannelGet = "get"; //  "com.k2.microsoft.teams.channel.get";
ChannelList = "list"; //  "com.k2.microsoft.teams.channel.list";
ChannelCreate = "create"; //  "com.k2.microsoft.teams.channel.create";
ChannelDelete = "delete"; //  "com.k2.microsoft.teams.channel.delete";
ChannelUpdate = "update"; //  "com.k2.microsoft.teams.channel.update";
ChannelSendMessage = "sendMessage"; //  "com.k2.microsoft.teams.channel.sendmessage";

//
// Tab
TabId = "id"; //  "com.k2.microsoft.teams.tab.id";
TabDisplayName = "displayName"; //  "com.k2.microsoft.teams.tab.displayname";
TabWeburl = "weburl"; //  "com.k2.microsoft.teams.tab.weburl";
TabConfigEntityId = "configEntityId"; //  "com.k2.microsoft.teams.tab.config.entityid";
TabConfigContentUrl = "configContentUrl"; //  "com.k2.microsoft.teams.tab.config.contenturl";
TabConfigWebsiteUrl = "configWebsiteUrl"; //  "com.k2.microsoft.teams.tab.config.websiteurl";
TabConfigRemoveUrl = "configRemoveUrl"; //  "com.k2.microsoft.teams.tab.config.removeurl";
TabTeamsAppAppId = "teamsAppAppId"; //  "com.k2.microsoft.teams.tab.teamsapp.appid";
TabTeamsAppExternalId = "teamsAppExternalId"; //  "com.k2.microsoft.teams.tab.teamsapp.externalid";
TabTeamsAppAppDisplayName = "teamsAppAppDisplayName"; //  "com.k2.microsoft.teams.tab.teamsapp.appdisplayname";
TabTeamsAppDistMethod = "teamsAppDistMethod"; //  "com.k2.microsoft.teams.tab.teamsapp.distmethod";
TabSortOrderIndex = "sortOrderIndex"; //  "com.k2.microsoft.teams.tab.sortorderindex";
TabIsSuccessful = "isSuccessful"; //  "com.k2.microsoft.teams.tab.issuccessful";
TabTeamId = "teamId"; //  "com.k2.microsoft.teams.tab.teamid";
TabChannelId = "channelId"; //  "com.k2.microsoft.teams.tab.channelid";

TabGet = "get"; //  "com.k2.microsoft.teams.tab.get";
TabList = "list"; //  "com.k2.microsoft.teams.tab.list";
TabCreateWordTab = "createWordTab"; //  "com.k2.microsoft.teams.tab.createwordtab";
TabCreateExcelTab = "createExcelTab"; //  "com.k2.microsoft.teams.tab.createexceltab";
TabCreatePowerpointTab = "createPowerpointTab"; //  "com.k2.microsoft.teams.tab.createpowerpointtab";
TabCreatePdfTab = "createPdfTab"; //  "com.k2.microsoft.teams.tab.createpdftab";
TabCreateOneNoteTab = "createOneNoteTab"; //  "com.k2.microsoft.teams.tab.createonenotetab";
TabCreatePlannerTab = "createPlannerTab"; //  "com.k2.microsoft.teams.tab.createplannertab";
TabCreateSharePointTab = "creatSsharePointTab"; //  "com.k2.microsoft.teams.tab.createsharepointtab";
TabCreateMsFormsTab = "createMsFormsTab"; //  "com.k2.microsoft.teams.tab.createmsformstab";
TabCreateMsStreamTab = "createmsStreamTab"; //  "com.k2.microsoft.teams.tab.createmsstreamtab";
TabCreateWebsiteTab = "createWebsiteTab"; //  "com.k2.microsoft.teams.tab.createwebsitetab";
TabCreateWikiTab = "createWikiTab"; //  "com.k2.microsoft.teams.tab.createwikitab";
TabCreatePowerBiTab = "createPowerBiTab"; //  "com.k2.microsoft.teams.tab.createpowerbitab";
TabCreateDocumentLibraryTab = "createDocumentLibraryTab"; //  "com.k2.microsoft.teams.tab.createdocumentlibrarytab";
TabCreateCustomTab = "createCustomTab"; //  "com.k2.microsoft.teams.tab.createcustomtab";
TabDelete = "delete"; //  "com.k2.microsoft.teams.tab.delete";
TabUpdate = "update"; //  "com.k2.microsoft.teams.tab.update";

//
// Apps
AppsId = "id"; //  "com.k2.microsoft.teams.apps.id";
AppsTeamId = "teamId"; //  "com.k2.microsoft.teams.apps.teamid";
AppsDisplayName = "displayName"; //  "com.k2.microsoft.teams.apps.displayname";
AppsVersion = "version"; //  "com.k2.microsoft.teams.apps.version";
AppsTeamAppDefinitionId = "teamsAppDefinitionId"; //  "com.k2.microsoft.teams.apps.teamsappdefinitionid";
AppsTeamsAppId = "teamsAppId"; //  "com.k2.microsoft.teams.apps.teamsappid";

AppsList = "list"; //  "com.k2.microsoft.teams.apps.list";

//K2 Required
ondescribe = function() {
    postSchema({

        objects: {
            [Team]: {
                displayName: "Team",
                description: "Team",
                properties: {
                    [TeamId]: {
                        displayName: "Team Id",
                        type: "string"
                    },
                    [TeamWeburl]: {
                        displayName: "Web Url",
                        type: "string"
                    },
                    [TeamDisplayName]: {
                        displayName: "Display Name",
                        type: "string"
                    },
                    [TeamCreationDate]: {
                        displayName: "Created On",
                        type: "string"
                    },
                    [TeamDescription]: {
                        displayName: "Description",
                        type: "string"
                    },
                    [TeamEmail]: {
                        displayName: "Email",
                        type: "string"
                    },
                    [TeamMailEnabled]: {
                        displayName: "Mail Enabled",
                        type: "boolean"
                    },
                    [TeamMailNickname]: {
                        displayName: "Mail Nick Name",
                        type: "string"
                    },
                    [TeamArchiveStatus]: {
                        displayName: "Archive Status",
                        type: "string"
                    },
                    [TeamOperationId]: {
                        displayName: "Operation Id",
                        type: "string"
                    },
                    [TeamOperationType]: {
                        displayName: "Operation Type",
                        type: "string"
                    },
                    [TeamLastActionDate]: {
                        displayName: "Last Action Date",
                        type: "string"
                    },
                    [TeamAttemptsCount]: {
                        displayName: "Attempts Count",
                        type: "string"
                    },
                    [TeamTargetResourceId]: {
                        displayName: "Target Resource Id",
                        type: "string"
                    },
                    [TeamTargetResourceLocation]: {
                        displayName: "Target Resource Location",
                        type: "string"
                    },
                    [TeamArchiveError]: {
                        displayName: "Error",
                        type: "string"
                    },
                    [TeamUserPrincipalName]: {
                        displayName: "User Principal Name",
                        type: "string"
                    },
                    [TeamResourceProvisioningOptions]: {
                        displayName: "Resource Provisioning Options",
                        type: "string"
                    },
                    [TeamIsArchived]: {
                        displayName: "Is Archived",
                        type: "boolean"
                    },
                    [TeamIsSuccessful]: {
                        displayName: "Is Successful",
                        type: "boolean"
                    },
                    [TeamRequestStatusUrl]: {
                        displayName: "Request Status Url",
                        type: "boolean"
                    },
                    [TeamMsAllowCreateUpdateChannels]: {
                        displayName: "Allow create/update channels to members",
                        type: "boolean"
                    },
                    [TeamMsAllowDeleteChannels]: {
                        displayName: "Allow delete channels to members",
                        type: "boolean"
                    },
                    [TeamMsAllowAddRemoveApps]: {
                        displayName: "Allow add/remove apps to members",
                        type: "boolean"
                    },
                    [TeamMsAllowCreateUpdateRemoveTabs]: {
                        displayName: "Allow create/update/remove tabs to members",
                        type: "boolean"
                    },
                    [TeamMsAllowCreateUpdateRemoveConnectors]: {
                        displayName: "Allow create/update/remove connectors to members",
                        type: "boolean"
                    },
                    [TeamGsAllowCreateUpdateChannels]: {
                        displayName: "Allow create/update channels to guests",
                        type: "boolean"
                    },
                    [TeamGsAllowDeleteChannels]: {
                        displayName: "Allow delete channels to guests",
                        type: "boolean"
                    },
                    [TeamMsgAllowUserEditMessages]: {
                        displayName: "Allow user to edit messages",
                        type: "boolean"
                    },
                    [TeamMsgAllowUserDeleteMessages]: {
                        displayName: "Allow user to delete messages",
                        type: "boolean"
                    },
                    [TeamMsgAllowOwnerDeleteMessages]: {
                        displayName: "Allow owner to delete messages",
                        type: "boolean"
                    },
                    [TeamMsgAllowTeamMentions]: {
                        displayName: "Allow team mentions",
                        type: "boolean"
                    },
                    [TeamMsgAllowChannelMentions]: {
                        displayName: "Allow channel mentions",
                        type: "boolean"
                    },
                    [TeamFsAllowGiphy]: {
                        displayName: "Allow Giphy",
                        type: "boolean"
                    },
                    [TeamFsAllowStickersAndMemes]: {
                        displayName: "Allow stickers and memes",
                        type: "boolean"
                    },
                    [TeamFsAllowCustomMemes]: {
                        displayName: "Allow custom memes",
                        type: "boolean"
                    }
                },
                methods: {
                    [TeamGet]: {
                        displayName: "Get",
                        type: "read",
                        inputs: [TeamId],
                        requiredInputs: [TeamId],
                        outputs: [
                            TeamId,
                            TeamDisplayName,
                            TeamCreationDate,
                            TeamDescription,
                            TeamEmail,
                            TeamMailEnabled,
                            TeamMailNickname,
                            TeamWeburl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamCreate]: {
                        displayName: "Create",
                        description: "Creates a new group and adds a team to the group",
                        type: "create",
                        inputs: [
                            TeamDisplayName,
                            TeamDescription,
                            TeamMailEnabled,
                            TeamMailNickname,
                            TeamUserPrincipalName
                        ],
                        requiredInputs: [TeamDisplayName],
                        outputs: [
                            TeamId,
                            TeamDisplayName,
                            TeamCreationDate,
                            TeamDescription,
                            TeamEmail,
                            TeamMailEnabled,
                            TeamMailNickname,
                            TeamWeburl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamAdd]: {
                        displayName: "Add a team to Existing Group",
                        description: "Add a team to an already existing AAD Group",
                        type: "create",
                        inputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        outputs: [TeamId,
                            TeamDisplayName,
                            TeamCreationDate,
                            TeamDescription,
                            TeamEmail,
                            TeamMailEnabled,
                            TeamMailNickname,
                            TeamWeburl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamArchive]: {
                        displayName: "Archive",
                        description: "Archive Team",
                        type: "execute",
                        inputs: [TeamId],
                        requiredInputs: [TeamId],
                        outputs: [TeamId,
                            TeamRequestStatusUrl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamUnarchive]: {
                        displayName: "Unarchive",
                        description: "Unarchive Team",
                        type: "execute",
                        inputs: [TeamId],
                        requiredInputs: [TeamId],
                        outputs: [TeamId,
                            TeamRequestStatusUrl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamCheckStatus]: {
                        displayName: "Check Archival Status",
                        description: "Check the stauts of an archival job.",
                        type: "execute",
                        parameters: {
                            [TeamArchiveOperationUrl]: {
                                displayName: "Archive/Unarchive operation URL",
                                type: "string"
                            }
                        },
                        requiredParameters: [TeamArchiveOperationUrl],
                        outputs: [TeamOperationId,
                            TeamOperationType,
                            TeamCreationDate,
                            TeamArchiveStatus,
                            TeamLastActionDate,
                            TeamAttemptsCount,
                            TeamTargetResourceId,
                            TeamTargetResourceLocation,
                            TeamArchiveError
                        ]
                    },
                    [TeamAddMember]: {
                        displayName: "Add Member",
                        description: "Add member to a team/group",
                        type: "create",
                        inputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        requiredInputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        outputs: [TeamIsSuccessful
                        ]
                    },
                    [TeamUpdate]: {
                        displayName: "Update",
                        description: "Updates a team settings",
                        type: "update",
                        inputs: [TeamId,
                            TeamMsAllowCreateUpdateChannels,
                            TeamMsAllowDeleteChannels,
                            TeamMsAllowAddRemoveApps,
                            TeamMsAllowCreateUpdateRemoveTabs,
                            TeamMsAllowCreateUpdateRemoveConnectors,
                            TeamGsAllowCreateUpdateChannels,
                            TeamGsAllowDeleteChannels,
                            TeamMsgAllowUserEditMessages,
                            TeamMsgAllowUserDeleteMessages,
                            TeamMsgAllowUserDeleteMessages,
                            TeamMsgAllowTeamMentions,
                            TeamMsgAllowChannelMentions,
                            TeamFsAllowGiphy,
                            TeamFsAllowStickersAndMemes,
                            TeamFsAllowCustomMemes
                        ],
                        requiredInputs: [TeamId],
                        outputs: [TeamIsSuccessful
                        ]
                    },
                    [TeamClone]: {
                        displayName: "Clone",
                        description: "Makes a copy of an existing Team",
                        type: "create",
                        inputs: [TeamId,
                            TeamDisplayName,
                            TeamDescription,
                            TeamMailNickname
                        ],
                        requiredInputs: [TeamId,
                            TeamDisplayName,
                            TeamDescription,
                            TeamMailNickname
                        ],
                        outputs: [TeamId,
                            TeamRequestStatusUrl,
                            TeamIsSuccessful
                        ]
                    },
                    [TeamAddOwner]: {
                        displayName: "Add Owner",
                        description: "Add owner to a team/group",
                        type: "execute",
                        parameters: {
                            [TeamAddAsMemberAlso]: {
                                displayName: "Add as Member Also",
                                type: "boolean"
                            }
                        },
                        inputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        requiredInputs: [TeamId,
                            TeamUserPrincipalName
                        ],
                        outputs: [TeamIsSuccessful
                        ]
                    },
                    [TeamList]: {
                        displayName: "List all teams/groups",
                        description: "List all teams/groups",
                        type: "list",
                        parameters: {
                            [TeamDisplayNameStartsWith]: {
                                displayName: "Display Name Starts With",
                                type: "string"
                            }
                        },
                        outputs: [TeamId,
                            TeamDisplayName,
                            TeamResourceProvisioningOptions
                        ]
                    },
                    [TeamMyTeamsList]: {
                        displayName: "List my teams",
                        description: "List my teams",
                        type: "list",
                        outputs: [TeamId,
                            TeamDisplayName,
                            TeamDescription,
                            TeamIsArchived
                        ]
                    }
                }
            },
            [Channel]: {
                displayName: "Channel",
                description: "Channel",
                properties: {
                    [ChannelId]: {
                        displayName: "Channel Id",
                        description: "Channel Id",
                        type: "string"
                    },
                    [ChannelDisplayName]: {
                        displayName: "Display Name",
                        description: "Display Name",
                        type: "string"
                    },
                    [ChannelDescription]: {
                        displayName: "Description",
                        description: "Description",
                        type: "string"
                    },
                    [ChannelEmail]: {
                        displayName: "Email",
                        description: "Email",
                        type: "string"
                    },
                    [ChannelWeburl]: {
                        displayName: "Web URL",
                        description: "Web URL",
                        type: "string"
                    },
                    [ChannelIsSuccessful]: {
                        displayName: "Is Successful",
                        description: "Is Successful",
                        type: "string"
                    },
                    [ChannelTeamId]: {
                        displayName: "Team Id",
                        description: "Team Id",
                        type: "string"
                    },
                    [ChannelMessageSubject]: {
                        displayName: "Subject",
                        description: "Message Subject",
                        type: "string"
                    },
                    [ChannelMessageBody]: {
                        displayName: "Body",
                        description: "Message Body",
                        type: "string"
                    },
                    [ChannelMessageIsImportant]: {
                        displayName: "Is Important",
                        description: "Message Importance (Normal/ High)",
                        type: "boolean"
                    }
                },
                methods: {
                    [ChannelGet]: {
                        displayName: "Get Channel",
                        type: "read",
                        inputs: [ChannelId,
                            ChannelTeamId
                        ],
                        requiredInputs: [ChannelId,
                            ChannelTeamId],
                        outputs: [ChannelId,
                            ChannelDisplayName,
                            ChannelDescription,
                            ChannelEmail,
                            ChannelWeburl,
                            ChannelIsSuccessful
                        ]
                    },
                    [ChannelList]: {
                        displayName: "List Channels",
                        type: "list",
                        inputs: [ChannelTeamId],
                        requiredInputs: [ChannelTeamId],
                        outputs: [ChannelId,
                            ChannelDisplayName,
                            ChannelDescription,
                            ChannelEmail
                        ]
                    },
                    [ChannelCreate]: {
                        displayName: "Create Channel",
                        type: "create",
                        inputs: [ChannelTeamId,
                            ChannelDisplayName,
                            ChannelDescription,
                        ],
                        requiredInputs: [ChannelTeamId,
                            ChannelDisplayName
                        ],
                        outputs: [ChannelId,
                            ChannelDisplayName,
                            ChannelDescription,
                            ChannelEmail,
                            ChannelWeburl,
                            ChannelIsSuccessful
                        ]
                    },
                    [ChannelDelete]: {
                        displayName: "Delete Channel",
                        type: "delete",
                        inputs: [ChannelId,
                            ChannelTeamId
                        ],
                        requiredInputs: [ChannelId,
                            ChannelTeamId],
                        outputs: [
                            ChannelIsSuccessful
                        ]
                    },
                    [ChannelUpdate]: {
                        displayName: "Update Channel",
                        type: "update",
                        inputs: [ChannelTeamId,
                            ChannelId,
                            ChannelDisplayName,
                            ChannelDescription
                        ],
                        requiredInputs: [ChannelId,
                            ChannelTeamId,
                            ChannelDisplayName,
                            ChannelDescription],
                        outputs: [ChannelIsSuccessful]
                    },
                    [ChannelSendMessage]: {
                        displayName: "Send Message to a channel",
                        type: "create",
                        inputs: [ChannelTeamId,
                            ChannelId,
                            ChannelMessageSubject,
                            ChannelMessageBody,
                            ChannelMessageIsImportant
                        ],
                        requiredInputs: [ChannelTeamId,
                            ChannelId,
                            ChannelMessageBody
                        ],
                        outputs: [ChannelIsSuccessful]
                    }
                }
            },
            [Tab]: {
                displayName: "Tabs",
                description: "Tabs",
                properties: {
                    [TabId]: {
                        displayName: "Tab Id",
                        description: "Tab Id",
                        type: "string"
                    },
                    [TabDisplayName]: {
                        displayName: "Display Name",
                        description: "Display Name",
                        type: "string"
                    },
                    [TabConfigEntityId]: {
                        displayName: "Entity Id",
                        description: "Entity Id",
                        type: "string"
                    },
                    [TabConfigContentUrl]: {
                        displayName: "Content URL",
                        description: "Content URL",
                        type: "string"
                    },
                    [TabConfigWebsiteUrl]: {
                        displayName: "Website URL",
                        description: "Website URL",
                        type: "string"
                    },
                    [TabConfigRemoveUrl]: {
                        displayName: "Remove URL",
                        description: "Remove URL",
                        type: "string"
                    },
                    [TabTeamsAppAppId]: {
                        displayName: "App Id",
                        description: "App Id",
                        type: "string"
                    },
                    [TabTeamsAppExternalId]: {
                        displayName: "External Id",
                        description: "External Id",
                        type: "string"
                    },
                    [TabTeamsAppAppDisplayName]: {
                        displayName: "App Display Name",
                        description: "App Display Name",
                        type: "string"
                    },
                    [TabTeamsAppDistMethod]: {
                        displayName: "Distribution Method",
                        description: "Distribution Method",
                        type: "string"
                    },
                    [TabSortOrderIndex]: {
                        displayName: "Sort Order Index",
                        description: "Sort Order Index",
                        type: "string"
                    },
                    [TabWeburl]: {
                        displayName: "Web URL",
                        description: "Web URL",
                        type: "string"
                    },
                    [TabIsSuccessful]: {
                        displayName: "Is Successful",
                        description: "Is Successful",
                        type: "boolean"
                    },
                    [TabTeamId]: {
                        displayName: "Team Id",
                        description: "Team Id",
                        type: "string"
                    },
                    [TabChannelId]: {
                        displayName: "Channel Id",
                        description: "Channel Id",
                        type: "string"
                    }
                },
                methods: {
                    [TabGet]: {
                        displayName: "Get tab",
                        type: "read",
                        inputs: [TabId,
                            TabTeamId,
                            TabChannelId
                        ],
                        requiredInputs: [TabId,
                            TabTeamId,
                            TabChannelId
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabTeamsAppAppId,
                            TabTeamsAppExternalId,
                            TabTeamsAppAppDisplayName,
                            TabTeamsAppDistMethod,
                            TabSortOrderIndex,
                            TabWeburl
                        ]
                    },
                    [TabList]: {
                        displayName: "List tabs",
                        type: "list",
                        inputs: [TabTeamId,
                            TabChannelId
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl
                        ]
                    },
                    [TabCreateWordTab]: {
                        displayName: "Create MS Word tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateExcelTab]: {
                        displayName: "Create MS Excel tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePowerpointTab]: {
                        displayName: "Create MS Powerpoint tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePdfTab]: {
                        displayName: "Create PDF tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateOneNoteTab]: {
                        displayName: "Create One Note tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePlannerTab]: {
                        displayName: "Create Planner tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateSharePointTab]: {
                        displayName: "Create SharePoint tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateMsFormsTab]: {
                        displayName: "Create MS Forms tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateMsStreamTab]: {
                        displayName: "Create MS Stream tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateWebsiteTab]: {
                        displayName: "Create Website tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateWikiTab]: {
                        displayName: "Create Wiki tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreatePowerBiTab]: {
                        displayName: "Create PowerBI tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateDocumentLibraryTab]: {
                        displayName: "Create Document Library tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabCreateCustomTab]: {
                        displayName: "Create custom tab",
                        type: "create",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabTeamsAppAppId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabTeamsAppAppId,
                            TabDisplayName
                        ],
                        outputs: [TabId,
                            TabDisplayName,
                            TabWeburl,
                            TabConfigEntityId,
                            TabConfigContentUrl,
                            TabConfigWebsiteUrl,
                            TabConfigRemoveUrl,
                            TabIsSuccessful
                        ]
                    },
                    [TabDelete]: {
                        displayName: "Delete tab",
                        type: "delete",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabId
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabId
                        ],
                        outputs: [
                            TabIsSuccessful
                        ]
                    },
                    [TabUpdate]: {
                        displayName: "Update tab",
                        type: "update",
                        inputs: [TabTeamId,
                            TabChannelId,
                            TabId,
                            TabDisplayName
                        ],
                        requiredInputs: [TabTeamId,
                            TabChannelId,
                            TabId
                        ],
                        outputs: [TabIsSuccessful]
                    },
                }
            },
            [Apps]: {
                displayName: "Apps",
                description: "Apps",
                properties: {
                    [AppsId]: {
                        displayName: "App Id",
                        description: "App Id",
                        type: "string"
                    },
                    [AppsTeamId]: {
                        displayName: "Team Id",
                        description: "Team Id",
                        type: "string"
                    },
                    [AppsDisplayName]: {
                        displayName: "App Display Name",
                        description: "App Display Name",
                        type: "string"
                    },
                    [AppsVersion]: {
                        displayName: "version",
                        description: "version",
                        type: "string"
                    },
                    [AppsTeamAppDefinitionId]: {
                        displayName: "Teams Apps Definition Id",
                        description: "Teams Apps Definition Id",
                        type: "string"
                    },
                    [AppsTeamsAppId]: {
                        displayName: "Teams Apps Id",
                        description: "Teams Apps Id",
                        type: "string"
                    }
                },
                methods: {
                    [AppsList]: {
                        displayName: "List installed apps",
                        type: "list",
                        inputs: [AppsTeamId],
                        requiredInputs: [AppsTeamId],
                        outputs: [AppsId,
                            AppsDisplayName,
                            AppsVersion,
                            AppsTeamAppDefinitionId,
                            AppsTeamsAppId
                        ]
                    }
                }
            }
        }

    });
}

onexecute = function (objectName, methodName, parameters, properties) {
    switch (objectName) {
        case Team:
            onexecuteTeam(methodName, parameters, properties);
            break;
        case Channel:
            onexecuteChannel(methodName, parameters, properties);
            break;
        case Tab:
            onexecuteTab(methodName, parameters, properties);
            break;
        case Apps:
            onexecuteApp(methodName, parameters, properties);
            break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteApp(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case AppsList:
            onexecuteInstalledAppsList(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported..");
    }
}

function onexecuteTeam(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    try {
        parameters[TeamIsSuccessful] = true;
        switch (methodName) {
            case TeamGet:
                onexecuteTeamGet(parameters, properties);
                break;
            case TeamCreate:
                onexecuteTeamCreate(parameters, properties);
                break;
            case TeamAdd:
                onexecuteTeamAdd(parameters, properties);
                break;
            case TeamUpdate:
                onexecuteTeamUpdate(parameters, properties);
                break;
            case TeamList:
                onexecuteTeamList(parameters, properties);
                break;
            case TeamArchive:
                onexecuteTeamArchive(parameters, properties);
                break;
            case TeamUnarchive:
                onexecuteTeamUnarchive(parameters, properties);
                break;
            case TeamAddMember:
                onexecuteTeamAddMember(parameters, properties);
                break;
            case TeamClone:
                onexecuteTeamClone(parameters, properties);
                break;
            case TeamAddOwner:
                onexecuteTeamAddOwner(parameters, properties);
                break;
            case TeamMyTeamsList:
                onexecuteTeamMyTeamsList(parameters, properties);
                break;
            case TeamCheckStatus:
                onexecuteTeamCheckStatus(parameters, properties);
                break;
            default: throw new Error("The method " + methodName + " is not supported..");
        }
    }
    catch (errMsg) {
        // postResult({
        //     [TeamIsSuccessful]: false
        // });
        Promise.reject(errMsg);
    }
}

function onexecuteTab(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    try {
        parameters[TabIsSuccessful] = true;
        switch (methodName) {
            case TabGet:
                onexecuteTabGet(parameters, properties);
                break;
            case TabUpdate:
                onexecuteTabUpdate(parameters, properties);
                break;
            case TabList:
                onexecuteTabList(parameters, properties);
                break;
            case TabDelete:
                onexecuteTabDelete(parameters, properties);
                break;
            case TabCreateWordTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateExcelTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreatePowerpointTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreatePdfTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateOneNoteTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreatePlannerTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateSharePointTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateMsFormsTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateMsStreamTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateWebsiteTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateWikiTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreatePowerBiTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateDocumentLibraryTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            case TabCreateCustomTab:
                onexecuteTabCreate(methodName, parameters, properties);
                break;
            default: throw new Error("The method " + methodName + " is not supported..");
        }
    }
    catch (errMsg) {
        // postResult({
        //     [TabIsSuccessful]: false
        // });
        Promise.reject(errMsg);
    }
}

function onexecuteTeamGet(parameters: SingleRecord, properties: SingleRecord) {
        // outputs: [
        //     TeamId,
        //     TeamDisplayName,
        //     TeamCreationDate,
        //     TeamDescription,
        //     TeamEmail,
        //     TeamMailEnabled,
        //     TeamMailNickname,
        //     TeamWeburl,
        //     TeamIsSuccessful
        // ]
        parameters[TeamIsSuccessful] = true;
        parameters[TeamId] = properties[TeamId];
        // Get Group Details By Group ID
        GetGroupDetailsById(parameters, properties, function (b) {
            parameters[TeamDisplayName] = b.displayName;
            parameters[TeamCreationDate] = b.createdDateTime;
            parameters[TeamDescription] = b.description;
            parameters[TeamEmail] = b.mail;
            parameters[TeamMailEnabled] = b.mailEnabled;
            parameters[TeamMailNickname] = b.mailNickname;
            //Get Team Details By Group ID
            GetTeamDetailsByID(parameters, properties, function (c) {
                parameters[TeamWeburl] = c.webUrl;
                // console.log("hello new eroor...................");
                // throw new Error("hello new eroor...................");
                //Post Results
                CreateAndReturnTeamObject(parameters, properties);
            });
        });
}

function onexecuteTeamCreate(parameters: SingleRecord, properties: SingleRecord) {
    // Create a Group, then add a team and add its to it
    // outputs: [
    //     TeamId,
    //     TeamDisplayName,
    //     TeamCreationDate,
    //     TeamDescription,
    //     TeamEmail,
    //     TeamMailEnabled,
    //     TeamMailNickname,
    //     TeamWeburl,
    //     TeamIsSuccessful
    // ]
    parameters[TeamIsSuccessful] = true;
    CreateGroup(parameters, properties, function (a) {
        properties[TeamId] = parameters[TeamId] = a.id;

        parameters[TeamCreationDate] = a.createdDateTime;
        parameters[TeamDescription] = a.description;
        parameters[TeamDisplayName] = a.displayName;
        parameters[TeamEmail] = a.mail;
        parameters[TeamMailEnabled] = a.mailEnabled;
        parameters[TeamMailNickname] = a.mailNickname;
        //Create a Team under the above Group
        CreateTeam(parameters, properties, function (b) {
            parameters[TeamWeburl] = b.webUrl;
            //Get User
            // GetUser(parameters, properties, function (c) {
            //     parameters[TeamUserId] = c.id;
            //     //Add Group Owner
            //     AddGroupOwner(parameters, properties, function (d) {
            //         //Add Members to the Group
            //         AddGroupMembers(parameters, properties, function (d) {
            //             //Post Results
                         CreateAndReturnTeamObject(parameters, properties);
            //         });
            //     });
            // });
        });
    });
}

function GetGroupIdByMailNickName(parameters: SingleRecord, properties: SingleRecord, cb) {
    var component = encodeURIComponent(`?$filter=mailNickName&q='${properties[TeamMailNickname]}'`);
    var url = baseUriEndpoint + "/groups" + component;
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetGroupDetailsById(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/groups/" + properties[TeamId];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetTeamDetailsByID(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TeamId];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateGroup(parameters: SingleRecord, properties: SingleRecord, cb) {
    //Create Body for GROUP POST
    var data = JSON.stringify({
        "description": properties[TeamDescription],
        "displayName": properties[TeamDisplayName],
        "groupTypes": ["Unified"],
        "mailEnabled": properties[TeamMailEnabled],
        "mailNickname": properties[TeamMailNickname],
        "visibility":"Private",
        "securityEnabled": false
    });
    var url = baseUriEndpoint + "/groups/";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    // var data = JSON.stringify({
    //     "memberSettings": {
    //         "allowCreateUpdateChannels": true
    //     },
    //     "messagingSettings": {
    //         "allowUserEditMessages": true,
    //         "allowUserDeleteMessages": true
    //     },
    //     "funSettings": {
    //         "allowGiphy": true,
    //         "giphyContentRating": "strict"
    //     }
    // });
    var data = JSON.stringify({});
    var url = baseUriEndpoint + "/groups/" + properties[TeamId] + "/team";
    ExecuteRequest(url, data, "PUT", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function ArchiveTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "shouldSetSpoSiteReadOnlyForMembers": true
    });
    var url = baseUriEndpoint + "/teams/" + properties[TeamId] + "/archive";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function UnarchiveTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = null;
    var url = baseUriEndpoint + "/teams/" + properties[TeamId] + "/unarchive";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function UpdateTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    //TODO - define properties that has to be updated
    var data = JSON.stringify({
        "memberSettings": {
            "allowCreateUpdateChannels": properties[TeamMsAllowCreateUpdateChannels],
            "allowDeleteChannels": properties[TeamMsAllowDeleteChannels],
            "allowAddRemoveApps": properties[TeamMsAllowAddRemoveApps],
            "allowCreateUpdateRemoveTabs": properties[TeamMsAllowCreateUpdateRemoveTabs],
            "allowCreateUpdateRemoveConnectors": properties[TeamMsAllowCreateUpdateRemoveConnectors]
        },
        "guestSettings": {
            "allowCreateUpdateChannels": properties[TeamGsAllowCreateUpdateChannels],
            "allowDeleteChannels": properties[TeamGsAllowDeleteChannels]
        },
        "messagingSettings": {
            "allowUserEditMessages": properties[TeamMsgAllowUserEditMessages],
            "allowUserDeleteMessages": properties[TeamMsgAllowUserDeleteMessages],
            "allowOwnerDeleteMessages": properties[TeamMsgAllowUserDeleteMessages],
            "allowTeamMentions": properties[TeamMsgAllowTeamMentions],
            "allowChannelMentions": properties[TeamMsgAllowChannelMentions]
        },
        "funSettings": {
            "allowGiphy": properties[TeamFsAllowGiphy],
            "allowStickersAndMemes": properties[TeamFsAllowStickersAndMemes],
            "allowCustomMemes": properties[TeamFsAllowCustomMemes]
        }
    });
    var url = baseUriEndpoint + "/teams/" + properties[TeamId];
    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CloneTeam(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[TeamDisplayName],
        "description": properties[TeamDescription],
        "mailNickname": properties[TeamMailNickname],
        "partsToClone": "apps,tabs,settings,channels,members",
        "visibility": "public"
    });
    var url = baseUriEndpoint + "/teams/" + properties[TeamId] + "/clone";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetUser(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/users/" + properties[TeamUserPrincipalName];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function AddGroupOwner(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/users/" + parameters[TeamUserId]
    });
    var url = baseUriEndpoint + "/groups/" + properties[TeamId] + "/owners/$ref";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function AddGroupMembers(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/directoryObjects/" + parameters[TeamUserId]
    });
    var url = baseUriEndpoint + "/groups/" + properties[TeamId] + "/members/$ref";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function ExecuteRequest(url: string, data: string, requestType: string, cb) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4)
            return;
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
                TeamIsSuccessful: false
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
    if (String(properties[TeamId]).length > 0) {
        parameters[TeamId] = properties[TeamId];
    }

    postResult({
        [TeamId]: parameters[TeamId],
        [TeamDisplayName]: parameters[TeamDisplayName],
        [TeamCreationDate]: parameters[TeamCreationDate],
        [TeamDescription]: parameters[TeamDescription],
        [TeamEmail]: parameters[TeamEmail],
        [TeamMailEnabled]: parameters[TeamMailEnabled],
        [TeamMailNickname]: parameters[TeamMailNickname],
        [TeamWeburl]: parameters[TeamWeburl],
        [TeamIsSuccessful]: parameters[TeamIsSuccessful]
    });
}

function onexecuteTeamAdd(parameters: SingleRecord, properties: SingleRecord) {
    //TODO - Should we make a call to Get Group Details by ID to get the team object details - for returning back to K2 user
    // Add Team to a group
    CreateTeam(parameters, properties, function (b) {
        parameters[TeamWeburl] = b.webUrl;
        // Get user
        GetUser(parameters, properties, function (c) {
            parameters[TeamUserId] = c.id;
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
                [TeamIsSuccessful]: true
            });
        }
        //CreateAndReturnTeamObject(parameters, properties);
    });
}

function onexecuteTeamMyTeamsList(parameters: SingleRecord, properties: SingleRecord) {
    GetMyTeams(parameters, properties, function (a) {
        console.log(a);
        postResult(a.value.map(x => {
            return {
                [TeamId]: x.id,
                [TeamDisplayName]: x.displayName,
                [TeamDescription]: x.description,
                [TeamIsArchived]: x.isArchived
            };
        }));
    });
}

function onexecuteTeamList(parameters: SingleRecord, properties: SingleRecord) {
    GetTeams(parameters, properties, function (a) {
        console.log(a);
        postResult(a.value.map(x => {
            return {
                [TeamId]: x.id,
                [TeamDisplayName]: x.displayName,
                [TeamResourceProvisioningOptions]: x.resourceProvisioningOptions[0]
            };
        }));
    });
}

function GetTeams(parameters: SingleRecord, properties: SingleRecord, cb) {
    if (parameters[TeamDisplayNameStartsWith] == null || parameters[TeamDisplayNameStartsWith] == "") {
        let component = encodeURIComponent("?$select=id,displayName,resourceProvisioningOptions");
        var url = baseUriEndpoint + "/groups" + component;
    }
    else {
        let component = encodeURIComponent("?$filter=startswith(displayName, '" + parameters[TeamDisplayNameStartsWith] + "')&$select=id,displayName,resourceProvisioningOptions");
        var url = baseUriEndpoint + "/groups" + component;
    }
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetMyTeams(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/me/joinedTeams";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTeamArchive(parameters: SingleRecord, properties: SingleRecord) {
    ArchiveTeam(parameters, properties, function (b) {
        // CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: parameters[TeamRequestStatusUrl],
            [TeamIsSuccessful]: parameters[TeamIsSuccessful]
        });
    });
}

function onexecuteTeamUnarchive(parameters: SingleRecord, properties: SingleRecord) {
    UnarchiveTeam(parameters, properties, function (b) {
        CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: parameters[TeamRequestStatusUrl],
            [TeamIsSuccessful]: parameters[TeamIsSuccessful]
        });
    });
}

function CheckArchivalStatus(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = null;
    var url = baseUriEndpoint + "/" + parameters[TeamArchiveOperationUrl];
    ExecuteRequest(url, data, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTeamCheckStatus(parameters: SingleRecord, properties: SingleRecord) {
    CheckArchivalStatus(parameters, properties, function (b) {
        postResult({
            [TeamOperationId]: b.id,
            [TeamOperationType]: b.operationType,
            [TeamCreationDate]: b.createdDateTime,
            [TeamArchiveStatus]: b.status,
            [TeamLastActionDate]: b.lastActionDateTime,
            [TeamAttemptsCount]: b.attemptsCount,
            [TeamTargetResourceId]: b.targetResourceId,
            [TeamTargetResourceLocation]: b.targetResourceLocation,
            [TeamArchiveError]: b.error
        });
    });
}

function onexecuteTeamClone(parameters: SingleRecord, properties: SingleRecord) {
    CloneTeam(parameters, properties, function (b) {
        //CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: parameters[TeamRequestStatusUrl],
            [TeamIsSuccessful]: parameters[TeamIsSuccessful]
        });
    });
}

function onexecuteChannel(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    parameters[ChannelIsSuccessful] = true;
    switch (methodName) {
        case ChannelGet:
            onexecuteChannelGet(parameters, properties);
            break;
        case ChannelList:
            onexecuteChannelList(parameters, properties);
            break;
        case ChannelCreate:
            onexecuteChannelCreate(parameters, properties);
            break;
        case ChannelDelete:
            onexecuteChannelDelete(parameters, properties);
            break;
        case ChannelUpdate:
            onexecuteChannelUpdate(parameters, properties);
            break;
        case ChannelSendMessage:
            onexecuteSendMessage(parameters, properties);
            break;
        default: throw new Error("The channel method " + methodName + " is not supported...");
    }
}

function onexecuteTeamAddMember(parameters: SingleRecord, properties: SingleRecord) {
    GetUser(parameters, properties, function (b) {
        parameters[TeamUserPrincipalName] = b.userPrincipalName;
        parameters[TeamUserId] = b.id;
        AddGroupMembers(parameters, properties, function (c) {
            //ToDO - remove the if condition and handle in try catch block
            if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                postResult({
                    [TeamIsSuccessful]: true
                });
            }
        });
    });
}

function onexecuteTeamAddOwner(parameters: SingleRecord, properties: SingleRecord) {
    GetUser(parameters, properties, function (b) {
        parameters[TeamUserPrincipalName] = b.userPrincipalName;
        parameters[TeamUserId] = b.id;
        AddGroupOwner(parameters, properties, function (c) {
            var stringValue = String(parameters[TeamAddAsMemberAlso]);
            var boolValue = stringValue.toLowerCase() == 'true' ? true : false;
            if (boolValue) {
                AddGroupMembers(parameters, properties, function (d) {
                    if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                        postResult({
                            [TeamIsSuccessful]: true
                        });
                    }
                });
            }
            else {
                if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
                    postResult({
                        [TeamIsSuccessful]: true
                    });
                }
            }
        });
    });
}

function onexecuteChannelGet(parameters: SingleRecord, properties: SingleRecord) {
    GetChannel(parameters, properties, function (a) {
        parameters[ChannelId] = a.id;
        parameters[ChannelDisplayName] = a.displayName;
        parameters[ChannelDescription] = a.description;
        parameters[ChannelEmail] = a.email;
        parameters[ChannelWeburl] = a.webUrl;
        CreateAndReturnChannelObject(parameters, properties);
    });
}

function onexecuteChannelList(parameters: SingleRecord, properties: SingleRecord) {
    GetChannelList(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [ChannelId]: x.id,
                [ChannelDisplayName]: x.displayName,
                [ChannelDescription]: x.description,
                [ChannelEmail]: x.email
            };
        }));
    });
}

function onexecuteChannelCreate(parameters: SingleRecord, properties: SingleRecord) {
    CreateChannel(parameters, properties, function (a) {
        parameters[ChannelId] = a.id;
        parameters[ChannelDisplayName] = a.displayName;
        parameters[ChannelDescription] = a.description;
        parameters[ChannelEmail] = a.email;
        parameters[ChannelWeburl] = a.webUrl;
        CreateAndReturnChannelObject(parameters, properties);
    });
}

function onexecuteChannelUpdate(parameters: SingleRecord, properties: SingleRecord) {
    UpdateChannel(parameters, properties, function (a) {
        parameters[ChannelId] = a.id;
        parameters[ChannelDisplayName] = a.displayName;
        parameters[ChannelDescription] = a.description;
        parameters[ChannelEmail] = a.email;
        parameters[ChannelWeburl] = a.webUrl;
        // CreateAndReturnChannelObject(parameters, properties);
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}

function onexecuteChannelDelete(parameters: SingleRecord, properties: SingleRecord) {
    DeleteChannel(parameters, properties, function (a) {
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}

function DeleteChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels/" + properties[ChannelId];
    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[ChannelDisplayName],
        "description": properties[ChannelDescription],
    });
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function CreateAndReturnChannelObject(parameters: SingleRecord, properties: SingleRecord) {
    var ChannelId = String(parameters[ChannelId]);
    if (ChannelId == null || ChannelId == "undefined" || ChannelId == "" || ChannelId == undefined)
        parameters[ChannelId] = properties[ChannelId];
    postResult({
        [ChannelId]: parameters[ChannelId],
        [ChannelDisplayName]: parameters[ChannelDisplayName],
        [ChannelDescription]: parameters[ChannelDescription],
        [ChannelEmail]: parameters[ChannelEmail],
        [ChannelWeburl]: parameters[ChannelWeburl],
        [ChannelIsSuccessful]: parameters[ChannelIsSuccessful]
    });
}

function GetChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels/" + properties[ChannelId];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function GetChannelList(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels?$select=id, displayname, description, email";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function UpdateChannel(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[ChannelDisplayName],
        "description": properties[ChannelDescription],
        "email": "test@k2rocks.com"
    });
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels/" + properties[ChannelId];
    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteSendMessage(parameters: SingleRecord, properties: SingleRecord) {
    SendMessage(parameters, properties, function (a) {
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}

function SendMessage(parameters: SingleRecord, properties: SingleRecord, cb) {
    var importance = properties[ChannelMessageIsImportant] == "true" ? "High" : "Normal";
    var data = JSON.stringify({
        "subject": properties[ChannelMessageSubject],
        "importance": importance.toString(),
        "body": {
            "contentType": "html",
            "content": properties[ChannelMessageBody]
        }
    });
    var url = baseUriEndpointBeta + "/teams/" + properties[ChannelTeamId] + "/channels/" + properties[ChannelId] + "/messages";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}


function onexecuteTabGet(parameters: SingleRecord, properties: SingleRecord) {
    GetTabInformation(parameters, properties, function (a) {
        postResult({
            [TabId]: a.id,
            [TabDisplayName]: a.displayName,
            [TabConfigEntityId]: a.configuration.entityId,
            [TabConfigContentUrl]: a.configuration.contentUrl,
            [TabConfigWebsiteUrl]: a.configuration.websiteUrl,
            [TabConfigRemoveUrl]: a.configuration.removeUrl,
            [TabTeamsAppAppId]: a.teamsApp.id,
            [TabTeamsAppExternalId]: a.teamsApp.externalId,
            [TabTeamsAppAppDisplayName]: a.teamsApp.displayName,
            [TabTeamsAppDistMethod]: a.teamsApp.distributionMethod,
            [TabSortOrderIndex]: a.sortOrderIndex,
            [TabWeburl]: a.webUrl
        });
    });
}

function GetTabInformation(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/Channels/" + properties[TabChannelId] + "/tabs/" + properties[TabId] + "?$expand=teamsApp";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTabUpdate(parameters: SingleRecord, properties: SingleRecord) {
    UpdateTab(parameters, properties, function (a) {
        postResult({
            [TabIsSuccessful]: true
        });
    });
}

function UpdateTab(parameters: SingleRecord, properties: SingleRecord, cb) {
    var data = JSON.stringify({
        "displayName": properties[TabDisplayName]
    });
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/Channels/" + properties[TabChannelId] + "/tabs/" + properties[TabId];
    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function onexecuteTabList(parameters: SingleRecord, properties: SingleRecord) {
    GetTabList(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [TabId]: x.id,
                [TabDisplayName]: x.displayName,
                [TabWeburl]: x.webUrl
            };
        }));
    });
}

function onexecuteTabCreate(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName) {
        case TabCreateWordTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Word", properties));
            break;
        case TabCreateExcelTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Excel", properties));
            break;
        case TabCreatePowerpointTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Powerpoint", properties));
            break;
        case TabCreatePdfTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("PDF", properties));
            break;
        case TabCreateOneNoteTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("OneNote", properties));
            break;
        case TabCreatePlannerTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Planner", properties));
            break;
        case TabCreateSharePointTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("SharePoint", properties));
            break;
        case TabCreateMsFormsTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("MicrosoftForms", properties));
            break;
        case TabCreateMsStreamTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("MicrosoftStream", properties));
            break;
        case TabCreateWebsiteTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Website", properties));
            break;
        case TabCreateWikiTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Wiki", properties));
            break;
        case TabCreatePowerBiTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("PowerBI", properties));
            break;
        case TabCreateDocumentLibraryTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("DocumentLibrary", properties));
            break;
        case TabCreateCustomTab:
            prepareDataAndCreateTab(parameters, properties, getRequestBody("Custom", properties));
            break;
        default: throw new Error("The object " + methodName + " is not supported.");
    }
}


function prepareDataAndCreateTab(parameters: SingleRecord, properties: SingleRecord, requestBody: string) {
    CreateTab(parameters, properties, requestBody, function (a) {
        // CreateAndReturnChannelObject(parameters, properties);
        postResult({
            [TabId]: a.id,
            [TabDisplayName]: a.displayName,
            [TabWeburl]: a.webUrl,
            [TabConfigEntityId]: a.configuration.entityId,
            [TabConfigContentUrl]: a.configuration.contentUrl,
            [TabConfigWebsiteUrl]: a.configuration.websiteUrl,
            [TabConfigRemoveUrl]: a.configuration.removeUrl,
            [TabIsSuccessful]: true
        });
    });
}

function CreateTab(parameters: SingleRecord, properties: SingleRecord, data: string, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/channels/" + properties[TabChannelId] + "/tabs";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}

function getRequestBody(tabType: string, properties) {
    var data;
    switch (tabType) {
        case "Word":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.word"
            });
            break;
        case "Excel":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.excel"
            });
            break;
        case "Powerpoint":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.powerpoint"
            });
            break;
        case "PDF":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.file.staticviewer.pdf"
            });
            break;
        case "OneNote":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78"
            });
            break;
        case "Planner":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
            });
            break;
        case "SharePoint":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88"
            });
            break;
        case "MicrosoftForms":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/81fef3a6-72aa-4648-a763-de824aeafb7d"
            });
            break;
        case "MicrosoftStream":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoftstream.embed.skypeteamstab"
            });
            break;
        case "Website":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"
            });
            break;
        case "Wiki":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.wiki"
            });
            break;
        case "PowerBI":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.powerbi"
            });
            break;
        case "DocumentLibrary":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.files.sharepoint"
            });
            break;
        case "Custom":
            data = JSON.stringify({
                "displayName": properties[TabDisplayName],
                "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + properties[TabTeamsAppAppId]
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
                [TabIsSuccessful]: true
            });
        }
    });
}

function DeleteTab(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/Channels/" + properties[TabChannelId] + "/tabs/" + properties[TabId];
    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}


function GetInstalledAppsList(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[AppsTeamId] + "/installedApps?$expand=teamsAppDefinition";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}


function onexecuteInstalledAppsList(parameters: SingleRecord, properties: SingleRecord) {
    GetInstalledAppsList(parameters, properties, function (a) {
        postResult(a.value.map(x => {
            return {
                [AppsId]: x.id,
                [AppsDisplayName]: x.teamsAppDefinition.displayName,
                [AppsVersion]: x.teamsAppDefinition.version,
                [AppsTeamAppDefinitionId]: x.teamsAppDefinition.id,
                [AppsTeamsAppId]: x.teamsAppDefinition.teamsAppId
            };
        }));
    });
}

function GetTabList(parameters: SingleRecord, properties: SingleRecord, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/channels/" + properties[TabChannelId] + "/tabs?$select=id,displayName,webUrl";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}