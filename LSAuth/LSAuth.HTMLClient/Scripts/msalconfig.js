var msalconfig = {
    graphApiEnpoint: "https://graph.microsoft.com",
    graphMeEndpoint: "https://graph.microsoft.com/beta/me/sendMail",
    graphUsersEndpoint: "https://graph.microsoft.com/v1.0/users",
    graphUserEndpoint: "https://graph.microsoft.com/v1.0/user",
    graphInvitationEndpoint: "https://graph.microsoft.com/beta/invitations",
    graphGroupMembersEndpoint: "https://graph.microsoft.com/v1.0/groups/{groupId}/members/$ref",
    graphAPIScopes2: ["https://graph.microsoft.com/.default"],
    graphAPIScopes: ["https://graph.microsoft.com/mail.send", "39683401-a60c-4bc6-b744-6dbaedef498d"],
    authorityHostUrl: "https://login.microsoftonline.com/infoconsulting.onmicrosoft.com",
    graphClientEndpoint: "https://functionsendmail.azurewebsites.net/api/HttpTriggerJS1?code=DwYi1XJUuuSJyzHMMxc6fuGabSWeFOLjromNG3FPLhvnFLYlPg/Fng==",
    //graphClientEndpoint: "https://graphclient.azurewebsites.net/api/GraphClient?code=MT16aSrNGvwV5AxBnLPI04kutOT1vlMB9mIAejHZDvm6cvfSbtVcAw==",
    common: "common/oauth2/v2.0",
    tenant: "infoconsulting.onmicrosoft.com",
    //clientID: "39683401-a60c-4bc6-b744-6dbaedef498d",
    clientID: "ea6cd5ab-9c40-4e27-b86b-8aaf86bae36a",
    secret: "",
    redirectUri2: "https://localhost:44300/",
    redirectUri: "http://localhost:57245/HTMLClient/default.htm",
    groupId: "c3346f2f-94d5-401a-afb1-9ad77616d79f"
};

/* ADAL config
*  @property {string} tenant - Your target tenant.
*  @property {string} clientId - Client ID assigned to your app by Azure Active Directory.
*  @property {string} redirectUri - Endpoint at which you expect to receive tokens.Defaults to `window.location.href`.
*  @property {string} instance - Azure Active Directory Instance.Defaults to `https://login.microsoftonline.com/`.
*  @property {Array} endpoints - Collection of {Endpoint-ResourceId} used for automatically attaching tokens in webApi calls.
*  @property {Boolean} popUp - Set this to true to enable login in a popup winodow instead of a full redirect.Defaults to `false`.
*  @property {string} localLoginUrl - Set this to redirect the user to a custom login page.
*  @property {function} displayCall - User defined function of handling the navigation to Azure AD authorization endpoint in case of login. Defaults to 'null'.
*  @property {string} postLogoutRedirectUri - Redirects the user to postLogoutRedirectUri after logout. Defaults is 'redirectUri'.
*  @property {string} cacheLocation - Sets browser storage to either 'localStorage' or sessionStorage'. Defaults to 'sessionStorage'.
*  @property {Array.<string>} anonymousEndpoints Array of keywords or URI's. Adal will not attach a token to outgoing requests that have these keywords or uri. Defaults to 'null'.
*  @property {number} expireOffsetSeconds If the cached token is about to be expired in the expireOffsetSeconds (in seconds), Adal will renew the token instead of using the cached token. Defaults to 120 seconds.
*  @property {string} correlationId Unique identifier used to map the request with the response. Defaults to RFC4122 version 4 guid (128 bits).
*  @property {number} loadFrameTimeout The number of milliseconds of inactivity before a token renewal response from AAD should be considered timed out.
*/
var adalconfig = {
    tenant: "infoconsulting.onmicrosoft.com",
    //clientId: "39683401-a60c-4bc6-b744-6dbaedef498d",
    clientId: "ea6cd5ab-9c40-4e27-b86b-8aaf86bae36a",
    redirectUri: "http://localhost:57245/HTMLClient/login.htm",
    //instance: "https://login.microsoftonline.com/infoconsulting.onmicrosoft.com"
}
/*
addUser(name)
addGroupMember(groupId, userId)
inviteUser(name, email)
getTokenThenCall(callback)
callGraphApi(token, endpoint, method, body, options)
*/