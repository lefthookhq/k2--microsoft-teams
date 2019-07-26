(function () {var b="https://graph.microsoft.com/v1.0";function c(e,t,a){switch(e){case"com.k2.microsoft.teams.team.url.get":d(t,a);break;case"com.k2.microsoft.teams.team.get":f(t,a);break;default:throw new Error("The method "+e+" is not supported.");}}function d(e,t){var a=new XMLHttpRequest;a.onreadystatechange=function(){if(4===a.readyState){if(200!==a.status)throw new Error("Failed with status "+a.status);console.log(a.responseText);var e=JSON.parse(a.responseText);postResult({"com.k2.microsoft.teams.team.id":e.id,"com.k2.microsoft.teams.team.weburl":e.webUrl,"com.k2.microsoft.teams.team.isarchived":e.isArchived})}},a.withCredentials=!0,a.open("GET",b+"/teams/"+e["com.k2.microsoft.teams.team.teamid"]),a.send()}function f(e,t){var a=new XMLHttpRequest;a.onreadystatechange=function(){if(4===a.readyState){if(200!==a.status)throw new Error("Failed with status "+a.status);console.log(a.responseText);var e=JSON.parse(a.responseText);postResult({"com.k2.microsoft.teams.team.id":e.id,"com.k2.microsoft.teams.team.displayname":e.displayName,"com.k2.microsoft.teams.team.creationdate":e.createdDateTime,"com.k2.microsoft.teams.team.description":e.description,"com.k2.microsoft.teams.team.email":e.mail,"com.k2.microsoft.teams.team.mailenabled":e.mailEnabled,"com.k2.microsoft.teams.team.mailnickname":e.mailNickname})}},a.withCredentials=!0,a.open("GET",b+"/groups/"+e["com.k2.microsoft.teams.team.teamid"]),a.send()}ondescribe=function(){postSchema({"com.k2.microsoft.teams":{displayName:"Microsoft Teams",description:"A service for integrating with Microsoft Teams.",objects:{"com.k2.microsoft.teams.team":{displayName:"Team",description:"Team",properties:{"com.k2.microsoft.teams.team.id":{displayName:"Id",type:"string"},"com.k2.microsoft.teams.team.weburl":{displayName:"Web Url",type:"string"},"com.k2.microsoft.teams.team.displayname":{displayName:"Display Name",type:"string"},"com.k2.microsoft.teams.team.creationdate":{displayName:"Created On",type:"dateTime"},"com.k2.microsoft.teams.team.description":{displayName:"Description",type:"string"},"com.k2.microsoft.teams.team.email":{displayName:"Email",type:"string"},"com.k2.microsoft.teams.team.mailenabled":{displayName:"Mail Enabled",type:"boolean"},"com.k2.microsoft.teams.team.mailnickname":{displayName:"Mail Nick Name",type:"string"}},methods:{"com.k2.microsoft.teams.team.url.get":{displayName:"Get Team Url",type:"read",parameters:{"com.k2.microsoft.teams.team.teamid":{displayName:"Team Id",type:"string"}},requiredParameters:["com.k2.microsoft.teams.team.teamid"],outputs:["com.k2.microsoft.teams.team.id","com.k2.microsoft.teams.team.weburl"]},"com.k2.microsoft.teams.team.get":{displayName:"Get Team",type:"read",parameters:{"com.k2.microsoft.teams.team.teamid":{displayName:"Team Id",type:"string"}},requiredParameters:["com.k2.microsoft.teams.team.teamid"],outputs:["com.k2.microsoft.teams.team.id","com.k2.microsoft.teams.team.displayname","com.k2.microsoft.teams.team.creationdate","com.k2.microsoft.teams.team.description","com.k2.microsoft.teams.team.email","com.k2.microsoft.teams.team.mailenabled","com.k2.microsoft.teams.team.mailnickname"]}}}}}})},onexecute=function(e,t,a,m){switch(e){case"com.k2.microsoft.teams.team":c(t,a,m);break;default:throw new Error("The object "+e+" is not supported.");}};})();