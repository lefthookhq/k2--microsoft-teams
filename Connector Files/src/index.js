import '@k2oss/k2-broker-core';

var metadata = {
    systemName: "CalebTest2",
    displayName: "CalebTest2 Example Broker",
    description: "An example broker"
};

// Constants used in place of service keys.
const baseUriEndpoint = "https://graph.microsoft.com/v1.0";
const baseUriEndpointBeta = "https://graph.microsoft.com/beta";

const Teams = "teams"; // "com.k2.microsoft.teams";
const Team = "team"; //  "com.k2.microsoft.teams.team";
const Channel = "channel"; //  "com.k2.microsoft.teams.channel";
const Tab = "tab"; //  "com.k2.microsoft.teams.tab";
const Apps = "apps"; //  "com.k2.microsoft.teams.apps";

//
// Team
const TeamId = "id"; //  "com.k2.microsoft.teams.team.id";
const TeamWeburl = "weburl"; //  "com.k2.microsoft.teams.team.weburl";
const TeamDisplayName = "displayName"; //  "com.k2.microsoft.teams.team.displayname";
const TeamCreationDate = "creationDate"; //  "com.k2.microsoft.teams.team.creationdate";
const TeamDescription = "description"; //  "com.k2.microsoft.teams.team.description";
const TeamEmail = "email"; //  "com.k2.microsoft.teams.team.email";
const TeamMailEnabled = "mailEnabled"; //  "com.k2.microsoft.teams.team.mailenabled";
const TeamMailNickname = "mailNickname"; //  "com.k2.microsoft.teams.team.mailnickname";
const TeamArchiveStatus = "archiveStatus"; //  "com.k2.microsoft.teams.team.archivestatus";
const TeamOperationId = "operationId"; //  "com.k2.microsoft.teams.team.operationid";
const TeamOperationType = "operationType"; //  "com.k2.microsoft.teams.team.operationtype";
const TeamLastActionDate = "lastActionDate"; //  "com.k2.microsoft.teams.team.lastactiondate";
const TeamAttemptsCount = "attemptsCount"; //  "com.k2.microsoft.teams.team.attemptscount";
const TeamTargetResourceId = "targetResourceId"; //  "com.k2.microsoft.teams.team.targetresourceid";
const TeamTargetResourceLocation = "targetResourceLocation"; //  "com.k2.microsoft.teams.team.targetresourcelocation";
const TeamArchiveError = "archiveError"; //  "com.k2.microsoft.teams.team.archiveerror";
const TeamUserPrincipalName = "userPrincipalName"; //  "com.k2.microsoft.teams.team.userprincipalname";
const TeamResourceProvisioningOptions = "resourceProvisioningOptions"; //  "com.k2.microsoft.teams.team.resourceprovisioningoptions";
const TeamIsArchived = "isArchived"; //  "com.k2.microsoft.teams.team.isarchived";
const TeamIsSuccessful = "isSuccessful"; //  "com.k2.microsoft.teams.team.issuccessful"
const TeamRequestStatusUrl = "requestStatusUrl"; //  "com.k2.microsoft.teams.team.requeststatusurl";
const TeamMsAllowCreateUpdateChannels = "msAllowCreateUpdateChannels"; //  "com.k2.microsoft.teams.team.msallowcreateupdatechannels";
const TeamMsAllowDeleteChannels = "msAllowDeleteChannels"; //  "com.k2.microsoft.teams.team.msallowdeletechannels";
const TeamMsAllowAddRemoveApps = "msAllowAddRemoveApps"; //  "com.k2.microsoft.teams.team.msallowaddremoveapps";
const TeamMsAllowCreateUpdateRemoveTabs = "msAllowCreateUpdateRemoveTabs"; //  "com.k2.microsoft.teams.team.msallowcreateupdateremovetabs";
const TeamMsAllowCreateUpdateRemoveConnectors = "msAllowCreateUpdateRemoveConnectors"; //  "com.k2.microsoft.teams.team.msallowcreateupdateremoveconnectors";
const TeamGsAllowCreateUpdateChannels = "gsAllowCreateUpdateChannels"; //  "com.k2.microsoft.teams.team.gsallowcreateupdatechannels";
const TeamGsAllowDeleteChannels = "gsAllowDeleteChannels"; //  "com.k2.microsoft.teams.team.gsallowdeletechannels";
const TeamMsgAllowUserEditMessages = "msgsAllowUserEditMessages"; //  "com.k2.microsoft.teams.team.msgsallowusereditmessages";
const TeamMsgAllowUserDeleteMessages = "msgsAllowUserDeleteMessages"; //  "com.k2.microsoft.teams.team.msgsallowuserdeletemessages";
const TeamMsgAllowOwnerDeleteMessages = "msgsAllowOwnerDeleteMessages"; //  "com.k2.microsoft.teams.team.msgsallowownerdeletemessages";
const TeamMsgAllowTeamMentions = "msgsAllowTeamMentions"; //  "com.k2.microsoft.teams.team.msgsallowteammentions";
const TeamMsgAllowChannelMentions = "msgsAllowChannelMentions"; //  "com.k2.microsoft.teams.team.msgsallowchannelmentions";
const TeamFsAllowGiphy = "fsAllowGiphy"; //  "com.k2.microsoft.teams.team.fsallowgiphy";
const TeamFsAllowStickersAndMemes = "fsAllowStickersAndMemes"; //  "com.k2.microsoft.teams.team.fsallowstickersandmemes";
const TeamFsAllowCustomMemes = "fsAllowCustomMemes"; //  "com.k2.microsoft.teams.team.fsallowcustommemes";

const TeamGet = "get"; //  "com.k2.microsoft.teams.team.get";
const TeamCreate = "create"; //  "com.k2.microsoft.teams.team.create";
const TeamAdd = "add"; //  "com.k2.microsoft.teams.team.add";
const TeamArchive = "archive"; //  "com.k2.microsoft.teams.team.archive";
const TeamUnarchive = "unarchive"; //  "com.k2.microsoft.teams.team.unarchive";
const TeamCheckStatus = "checkStatus"; //  "com.k2.microsoft.teams.team.checkstatus";
const TeamAddMember = "addMember"; //  "com.k2.microsoft.teams.team.addmember";
const TeamUpdate = "update"; //  "com.k2.microsoft.teams.team.update";
const TeamClone = "clone"; //  "com.k2.microsoft.teams.team.clone";
const TeamAddOwner = "addowner"; //  "com.k2.microsoft.teams.team.addowner";
const TeamList = "list"; //  "com.k2.microsoft.teams.team.list";
const TeamMyTeamsList = "myTeamsList"; //  "com.k2.microsoft.teams.team.myteamslist";

const TeamArchiveOperationUrl = "archiveOperationUrl"; //  "com.k2.microsoft.teams.team.archiveoperationurl";
const TeamAddAsMemberAlso = "addAsMemberAlso"; //  "com.k2.microsoft.teams.team.addasmemberalso";
const TeamDisplayNameStartsWith = "displayNameStartsWith"; //  "com.k2.microsoft.teams.team.displaynamestartswith";
const TeamUserId = "userId"; //  "com.k2.microsoft.teams.team.userid";

//
// Channel
const ChannelId = "id"; //  "com.k2.microsoft.teams.channel.id";
const ChannelDisplayName = "displayName"; //  "com.k2.microsoft.teams.channel.displayname";
const ChannelDescription = "description"; //  "com.k2.microsoft.teams.channel.description";
const ChannelEmail = "email"; //  "com.k2.microsoft.teams.channel.email";
const ChannelWeburl = "weburl"; //  "com.k2.microsoft.teams.channel.weburl";
const ChannelIsSuccessful = "isSuccessful"; //  "com.k2.microsoft.teams.channel.issuccessful";
const ChannelTeamId = "teamId"; //  "com.k2.microsoft.teams.channel.teamid";
const ChannelMessageSubject = "messageSubject"; //  "com.k2.microsoft.teams.channel.message.subject";
const ChannelMessageBody = "messageBody"; //  "com.k2.microsoft.teams.channel.message.body";
const ChannelMessageIsImportant = "messageIsImportant"; //  "com.k2.microsoft.teams.channel.message.isimportant";

const ChannelGet = "get"; //  "com.k2.microsoft.teams.channel.get";
const ChannelList = "list"; //  "com.k2.microsoft.teams.channel.list";
const ChannelCreate = "create"; //  "com.k2.microsoft.teams.channel.create";
const ChannelDelete = "delete"; //  "com.k2.microsoft.teams.channel.delete";
const ChannelUpdate = "update"; //  "com.k2.microsoft.teams.channel.update";
const ChannelSendMessage = "sendMessage"; //  "com.k2.microsoft.teams.channel.sendmessage";

//
// Tab
const TabId = "id"; //  "com.k2.microsoft.teams.tab.id";
const TabDisplayName = "displayName"; //  "com.k2.microsoft.teams.tab.displayname";
const TabWeburl = "weburl"; //  "com.k2.microsoft.teams.tab.weburl";
const TabConfigEntityId = "configEntityId"; //  "com.k2.microsoft.teams.tab.config.entityid";
const TabConfigContentUrl = "configContentUrl"; //  "com.k2.microsoft.teams.tab.config.contenturl";
const TabConfigWebsiteUrl = "configWebsiteUrl"; //  "com.k2.microsoft.teams.tab.config.websiteurl";
const TabConfigRemoveUrl = "configRemoveUrl"; //  "com.k2.microsoft.teams.tab.config.removeurl";
const TabTeamsAppAppId = "teamsAppAppId"; //  "com.k2.microsoft.teams.tab.teamsapp.appid";
const TabTeamsAppExternalId = "teamsAppExternalId"; //  "com.k2.microsoft.teams.tab.teamsapp.externalid";
const TabTeamsAppAppDisplayName = "teamsAppAppDisplayName"; //  "com.k2.microsoft.teams.tab.teamsapp.appdisplayname";
const TabTeamsAppDistMethod = "teamsAppDistMethod"; //  "com.k2.microsoft.teams.tab.teamsapp.distmethod";
const TabSortOrderIndex = "sortOrderIndex"; //  "com.k2.microsoft.teams.tab.sortorderindex";
const TabIsSuccessful = "isSuccessful"; //  "com.k2.microsoft.teams.tab.issuccessful";
const TabTeamId = "teamId"; //  "com.k2.microsoft.teams.tab.teamid";
const TabChannelId = "channelId"; //  "com.k2.microsoft.teams.tab.channelid";

const TabGet = "get"; //  "com.k2.microsoft.teams.tab.get";
const TabList = "list"; //  "com.k2.microsoft.teams.tab.list";
const TabCreateWordTab = "createWordTab"; //  "com.k2.microsoft.teams.tab.createwordtab";
const TabCreateExcelTab = "createExcelTab"; //  "com.k2.microsoft.teams.tab.createexceltab";
const TabCreatePowerpointTab = "createPowerpointTab"; //  "com.k2.microsoft.teams.tab.createpowerpointtab";
const TabCreatePdfTab = "createPdfTab"; //  "com.k2.microsoft.teams.tab.createpdftab";
const TabCreateOneNoteTab = "createOneNoteTab"; //  "com.k2.microsoft.teams.tab.createonenotetab";
const TabCreatePlannerTab = "createPlannerTab"; //  "com.k2.microsoft.teams.tab.createplannertab";
const TabCreateSharePointTab = "creatSsharePointTab"; //  "com.k2.microsoft.teams.tab.createsharepointtab";
const TabCreateMsFormsTab = "createMsFormsTab"; //  "com.k2.microsoft.teams.tab.createmsformstab";
const TabCreateMsStreamTab = "createmsStreamTab"; //  "com.k2.microsoft.teams.tab.createmsstreamtab";
const TabCreateWebsiteTab = "createWebsiteTab"; //  "com.k2.microsoft.teams.tab.createwebsitetab";
const TabCreateWikiTab = "createWikiTab"; //  "com.k2.microsoft.teams.tab.createwikitab";
const TabCreatePowerBiTab = "createPowerBiTab"; //  "com.k2.microsoft.teams.tab.createpowerbitab";
const TabCreateDocumentLibraryTab = "createDocumentLibraryTab"; //  "com.k2.microsoft.teams.tab.createdocumentlibrarytab";
const TabCreateCustomTab = "createCustomTab"; //  "com.k2.microsoft.teams.tab.createcustomtab";
const TabDelete = "delete"; //  "com.k2.microsoft.teams.tab.delete";
const TabUpdate = "update"; //  "com.k2.microsoft.teams.tab.update";

//
// Apps
const AppsId = "id"; //  "com.k2.microsoft.teams.apps.id";
const AppsTeamId = "teamId"; //  "com.k2.microsoft.teams.apps.teamid";
const AppsDisplayName = "displayName"; //  "com.k2.microsoft.teams.apps.displayname";
const AppsVersion = "version"; //  "com.k2.microsoft.teams.apps.version";
const AppsTeamAppDefinitionId = "teamsAppDefinitionId"; //  "com.k2.microsoft.teams.apps.teamsappdefinitionid";
const AppsTeamsAppId = "teamsAppId"; //  "com.k2.microsoft.teams.apps.teamsappid";

const AppsList = "list"; //  "com.k2.microsoft.teams.apps.list";

//K2 Required
export async function ondescribe() {
    postSchema({
        [Teams]: {
            displayName: "Microsoft Teams",
            description: "A service for integrating with Microsoft Teams.",
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
                            type: "dateTime"
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
                            type: "dateTime"
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
                        [TeamCreate]: {
                            displayName: "Create",
                            description: "Creates a new group and adds a team to the group",
                            type: "create",
                            inputs: [TeamDisplayName,
                                TeamDescription,
                                TeamMailEnabled,
                                TeamMailNickname,
                                TeamUserPrincipalName
                            ],
                            requiredInputs: [TeamDisplayName],
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
        }
    });
};
export async function onexecute(objectName, methodName, parameters, properties) {
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
};
// Begin standard JS --
function onexecuteApp(methodName, parameters, properties) {
    switch (methodName) {
        case AppsList:
            onexecuteInstalledAppsList(parameters, properties);
            break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}
function onexecuteTeam(methodName, parameters, properties) {
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
            default: throw new Error("The method " + methodName + " is not supported.");
        }
    }
    catch (errMsg) {
        postResult({
            [TeamIsSuccessful]: false
        });
    }
}
function onexecuteTab(methodName, parameters, properties) {
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
            default: throw new Error("The method " + methodName + " is not supported.");
        }
    }
    catch (errMsg) {
        postResult({
            [TabIsSuccessful]: false
        });
    }
}
function onexecuteTeamGet(parameters, properties) {
    try {
        parameters[TeamIsSuccessful] = true;
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
    catch (errMsg) {
        postResult({
            [TeamIsSuccessful]: false
        });
    }
}
function onexecuteTeamCreate(parameters, properties) {
    //Create a Group
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
            GetUser(parameters, properties, function (c) {
                parameters[TeamUserId] = c.id;
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
function GetGroupIdByMailNickName(parameters, properties, cb) {
    var url = baseUriEndpoint + "/groups?$filter=mailNickName%20eq%20'" + String(properties[TeamMailNickname]) + "'";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function GetGroupDetailsById(parameters, properties, cb) {
    var url = baseUriEndpoint + "/groups/" + properties[TeamId];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function GetTeamDetailsByID(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TeamId];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function CreateGroup(parameters, properties, cb) {
    //Create Body for GROUP POST
    var data = JSON.stringify({
        "description": properties[TeamDescription],
        "displayName": properties[TeamDisplayName],
        "groupTypes": ["Unified"],
        "mailEnabled": properties[TeamMailEnabled],
        "mailNickname": properties[TeamMailNickname],
        "securityEnabled": true
    });
    var url = baseUriEndpoint + "/groups/";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function CreateTeam(parameters, properties, cb) {
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
    var url = baseUriEndpoint + "/groups/" + properties[TeamId] + "/team";
    ExecuteRequest(url, data, "PUT", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function ArchiveTeam(parameters, properties, cb) {
    var data = JSON.stringify({
        "shouldSetSpoSiteReadOnlyForMembers": true
    });
    var url = baseUriEndpoint + "/teams/" + properties[TeamId] + "/archive";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function UnarchiveTeam(parameters, properties, cb) {
    var data = null;
    var url = baseUriEndpoint + "/teams/" + properties[TeamId] + "/unarchive";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function UpdateTeam(parameters, properties, cb) {
    //ToDo - define properties that has to be updated
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
function CloneTeam(parameters, properties, cb) {
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
function GetUser(parameters, properties, cb) {
    var url = baseUriEndpoint + "/users/" + properties[TeamUserPrincipalName];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function AddGroupOwner(parameters, properties, cb) {
    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/users/" + parameters[TeamUserId]
    });
    var url = baseUriEndpoint + "/groups/" + properties[TeamId] + "/owners/$ref";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function AddGroupMembers(parameters, properties, cb) {
    var data = JSON.stringify({
        "@odata.id": baseUriEndpoint + "/directoryObjects/" + parameters[TeamUserId]
    });
    var url = baseUriEndpoint + "/groups/" + properties[TeamId] + "/members/$ref";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function ExecuteRequest(url, data, requestType, cb) {
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
function CreateAndReturnTeamObject(parameters, properties) {
    if (String(properties[TeamId]).length > 0)
        parameters[TeamId] = properties[TeamId];
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
function onexecuteTeamAdd(parameters, properties) {
    //To Do - Should we make a call to Get Group Details by ID to get the team object details - for returning back to K2 user
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
function onexecuteTeamUpdate(parameters, properties) {
    UpdateTeam(parameters, properties, function (c) {
        if (c.responseText == null || c.responseText == "" || c.responseText == undefined || c.responseText == "undefined") {
            postResult({
                [TeamIsSuccessful]: true
            });
        }
        //CreateAndReturnTeamObject(parameters, properties);
    });
}
function onexecuteTeamMyTeamsList(parameters, properties) {
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
function onexecuteTeamList(parameters, properties) {
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
function GetTeams(parameters, properties, cb) {
    if (parameters[TeamDisplayNameStartsWith] == null || parameters[TeamDisplayNameStartsWith] == "")
        var url = baseUriEndpoint + "/groups?$select=id,displayName,resourceProvisioningOptions";
    else
        var url = baseUriEndpoint + "/groups?$filter=startswith(displayName, '" + parameters[TeamDisplayNameStartsWith] + "')&$select=id,displayName,resourceProvisioningOptions";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function GetMyTeams(parameters, properties, cb) {
    var url = baseUriEndpoint + "/me/joinedTeams";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function onexecuteTeamArchive(parameters, properties) {
    ArchiveTeam(parameters, properties, function (b) {
        // CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: parameters[TeamRequestStatusUrl],
            [TeamIsSuccessful]: parameters[TeamIsSuccessful]
        });
    });
}
function onexecuteTeamUnarchive(parameters, properties) {
    UnarchiveTeam(parameters, properties, function (b) {
        CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: parameters[TeamRequestStatusUrl],
            [TeamIsSuccessful]: parameters[TeamIsSuccessful]
        });
    });
}
function CheckArchivalStatus(parameters, properties, cb) {
    var data = null;
    var url = baseUriEndpoint + "/" + parameters[TeamArchiveOperationUrl];
    ExecuteRequest(url, data, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function onexecuteTeamCheckStatus(parameters, properties) {
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
function onexecuteTeamClone(parameters, properties) {
    CloneTeam(parameters, properties, function (b) {
        //CreateAndReturnTeamObject(parameters, properties);
        postResult({
            [TeamId]: properties[TeamId],
            [TeamRequestStatusUrl]: parameters[TeamRequestStatusUrl],
            [TeamIsSuccessful]: parameters[TeamIsSuccessful]
        });
    });
}
function onexecuteChannel(methodName, parameters, properties) {
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
function onexecuteTeamAddMember(parameters, properties) {
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
function onexecuteTeamAddOwner(parameters, properties) {
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
function onexecuteChannelGet(parameters, properties) {
    GetChannel(parameters, properties, function (a) {
        parameters[ChannelId] = a.id;
        parameters[ChannelDisplayName] = a.displayName;
        parameters[ChannelDescription] = a.description;
        parameters[ChannelEmail] = a.email;
        parameters[ChannelWeburl] = a.webUrl;
        CreateAndReturnChannelObject(parameters, properties);
    });
}
function onexecuteChannelList(parameters, properties) {
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
function onexecuteChannelCreate(parameters, properties) {
    CreateChannel(parameters, properties, function (a) {
        parameters[ChannelId] = a.id;
        parameters[ChannelDisplayName] = a.displayName;
        parameters[ChannelDescription] = a.description;
        parameters[ChannelEmail] = a.email;
        parameters[ChannelWeburl] = a.webUrl;
        CreateAndReturnChannelObject(parameters, properties);
    });
}
function onexecuteChannelUpdate(parameters, properties) {
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
function onexecuteChannelDelete(parameters, properties) {
    DeleteChannel(parameters, properties, function (a) {
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}
function DeleteChannel(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels/" + properties[ChannelId];
    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function CreateChannel(parameters, properties, cb) {
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
function CreateAndReturnChannelObject(parameters, properties) {
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
function GetChannel(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels/" + properties[ChannelId];
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function GetChannelList(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[ChannelTeamId] + "/channels?$select=id, displayname, description, email";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function UpdateChannel(parameters, properties, cb) {
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
function onexecuteSendMessage(parameters, properties) {
    SendMessage(parameters, properties, function (a) {
        postResult({
            [ChannelIsSuccessful]: true
        });
    });
}
function SendMessage(parameters, properties, cb) {
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
function onexecuteTabGet(parameters, properties) {
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
function GetTabInformation(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/Channels/" + properties[TabChannelId] + "/tabs/" + properties[TabId] + "?$expand=teamsApp";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function onexecuteTabUpdate(parameters, properties) {
    UpdateTab(parameters, properties, function (a) {
        postResult({
            [TabIsSuccessful]: true
        });
    });
}
function UpdateTab(parameters, properties, cb) {
    var data = JSON.stringify({
        "displayName": properties[TabDisplayName]
    });
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/Channels/" + properties[TabChannelId] + "/tabs/" + properties[TabId];
    ExecuteRequest(url, data, "PATCH", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function onexecuteTabList(parameters, properties) {
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
function onexecuteTabCreate(methodName, parameters, properties) {
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
function prepareDataAndCreateTab(parameters, properties, requestBody) {
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
function CreateTab(parameters, properties, data, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/channels/" + properties[TabChannelId] + "/tabs";
    ExecuteRequest(url, data, "POST", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function getRequestBody(tabType, properties) {
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
function onexecuteTabDelete(parameters, properties) {
    DeleteTab(parameters, properties, function (a) {
        if (a == null || a == "") {
            postResult({
                [TabIsSuccessful]: true
            });
        }
    });
}
function DeleteTab(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/Channels/" + properties[TabChannelId] + "/tabs/" + properties[TabId];
    ExecuteRequest(url, null, "DELETE", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function GetInstalledAppsList(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[AppsTeamId] + "/installedApps?$expand=teamsAppDefinition";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
function onexecuteInstalledAppsList(parameters, properties) {
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
function GetTabList(parameters, properties, cb) {
    var url = baseUriEndpoint + "/teams/" + properties[TabTeamId] + "/channels/" + properties[TabChannelId] + "/tabs?$select=id,displayName,webUrl";
    ExecuteRequest(url, null, "GET", function (responseText) {
        if (typeof cb === 'function')
            cb(responseText);
    });
}
