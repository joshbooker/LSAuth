var msalconfig = {
    graphUsersEndpoint: "https://graph.microsoft.com/v1.0/users",
    graphInvitationEndpoint: "https://graph.microsoft.com/beta/invitations",
    graphGroupMembersEndpoint: "https://graph.microsoft.com/v1.0/groups/{groupId}/members/$ref",
    graphAPIScopes: ["https://graph.microsoft.com/.default"],
    authorityHostUrl: "https://login.microsoftonline.com",
    common: "common/oauth2/v2.0",
    tenant: "infoconsulting.onmicrosoft.com",
    //authorityUrl: msalconfig.authorityHostUrl + '/' + msalconfig.common,
    clientID: "39683401-a60c-4bc6-b744-6dbaedef498d",
    secret: "",
    redirectUri2: "https://localhost:44300/",
    redirectUri: location.origin,
    groupId: "c3346f2f-94d5-401a-afb1-9ad77616d79f"
};

/*
addUser(name)
addGroupMember(groupId, userId)
inviteUser(name, email)
getTokenThenCall(callback)
callGraphApi(token, endpoint, method, body, options)
*/