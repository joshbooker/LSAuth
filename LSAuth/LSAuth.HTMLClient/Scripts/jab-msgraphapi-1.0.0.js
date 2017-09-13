(function (window, undefined) {

    /*
    addUser(name)
    addGroupMember(groupId, userId)
    inviteUser(name, email)
    callGraphApi(token, endpoint, method, body, options)
    getToken()
    isLoggedIn()
    login()
    logout()
    */

    /*
    TO DO:
        --handle more 400 response errors ie:  Invitee in Invited Tenant; Resource not Found; Password complexity error
    */
    var myGraph = window.graph || {};

    // AMD support
    if (typeof define === 'function' && define.amd) {
        define('myGraph', myGraph);
    } else {
        window.myGraph = myGraph;
    }

    myGraph.version = {
        major: 1,
        minor: 0,
        build: 0
    };

// Graph API endpoint to show user profile
var graphApiEndpoint = "https://graph.microsoft.com/v1.0/me";

//admin consent endpoint example
//https://login.microsoftonline.com/common/adminconsent?client_id=39683401-a60c-4bc6-b744-6dbaedef498d&state=12345&redirect_uri=http://localhost:30662/

// Graph API scope used to obtain the access token to read user profile
var graphAPIScopes = ["https://graph.microsoft.com/user.read"];
myGraph.graphAPIScopes = graphAPIScopes;

// Initialize application
var userAgentApplication = new Msal.UserAgentApplication(msalconfig.clientID, null, loginCallback, {
    redirectUri: msalconfig.redirectUri
});

//Previous version of msal uses redirect url via a property
if (userAgentApplication.redirectUri) {
    userAgentApplication.redirectUri = msalconfig.redirectUri;
}
myGraph.userAgentApplication = userAgentApplication;

myGraph.testCall = function callGraphApi2() {
    var result = myGraph.inviteUser("JoshTest",
                                "jbooker@midcoast.com",
                                msalconfig.redirectUri,
                                true,
                                "Hi Josh2, you can find more information here: " + msalconfig.redirectUri,
                                "josh@joshbooker.com",
                                "ccJosh")
                .then(function (data) {
                    return addGroupMember(msalconfig.groupId, data.invitedUser.id);
                });
    return result;
}

/*
POST https://graph.microsoft.com/v1.0/users
Content-type: application/json

{
    "accountEnabled": true,
    "displayName": "displayName-value",
    "mailNickname": "mailNickname-value",
    "userPrincipalName": "upn-value@tenant-value.onmicrosoft.com",
    "passwordProfile" : {
        "forceChangePasswordNextSignIn": true,
        "password": "password-value"
    }
}
*/

    /*
    * Add Internal User.
    * 
    * @param {any} displayName - displayName
    * @param {any} userPrincipalName - userPrincipalName
    * @param {any} mailNickname - mailNickname
    * @param {any} forceChangePasswordNextSignIn - forceChangePasswordNextSignIn
    * @param {any} password - password
    * @param {any} accountEnabled - accountEnabled
    */
    var addUser = function addUser(displayName, userPrincipalName, mailNickname, forceChangePasswordNextSignIn, password, accountEnabled) {
        var body = {
            "accountEnabled": accountEnabled,
            "displayName": displayName,
            "mailNickname": mailNickname,
            "userPrincipalName": userPrincipalName,
            "passwordProfile": {
                "forceChangePasswordNextSignIn": forceChangePasswordNextSignIn,
                "password": password
            }
        };
        return callGraphApi(msalconfig.graphUsersEndpoint, "POST", body)
                .then(function (response) {
                    //console.log("inviteUser Response: " + response);
                    var contentType = response.headers.get("content-type");
                    if ((response.status === 200 || response.status === 201) && contentType && contentType.indexOf("application/json") !== -1) {
                        return response.json()
                            .then(function (data) {
                                //console.log(data);
                                console.log("User added: " + data.displayName)
                                return data;
                            })
                            .catch(function (error) {
                                console.log("addUser json Error: " + error.message);
                                return error;
                            });
                    } else {
                        return response.json()
                            .then(function (data) {
                                console.log("addUser Failed: " + data.error.message);
                                return data;
                            })
                            .catch(function (error) {
                                console.log("addUser json2 Error: " + error.message);
                                return error;
                            });
                    }
                })
                .catch(function (error) {
                    console.log("addUser Error: " + error.message);
                    return error;
                });
    }
    myGraph.addUser = addUser;


    /*
    * Invite an External User.
    * 
    * @param {any} invitedUserDisplayName - invitedUserDisplayName
    * @param {any} invitedUserEmailAddress - invitedUserEmailAddress
    * @param {any} inviteRedirectUrl - inviteRedirectUrl
    * @param {any} sendInvitationMessage - sendInvitationMessage
    * @param {any} customizedMessageBody - customizedMessageBody
    * @param {any} ccRecipientEmailAddress - ccRecipientEmailAddress
    * @param {any} ccRecipientName - ccRecipientName
    */
    var inviteUser = function inviteUser(invitedUserDisplayName, invitedUserEmailAddress, inviteRedirectUrl, sendInvitationMessage, customizedMessageBody, ccRecipientEmailAddress, ccRecipientName) {
        var body = {
            "invitedUserDisplayName": invitedUserDisplayName,
            "invitedUserEmailAddress": invitedUserEmailAddress,
            "inviteRedirectUrl": inviteRedirectUrl,
            "sendInvitationMessage": sendInvitationMessage,
            "invitedUserMessageInfo": {
                "ccRecipients": [{
                    "emailAddress": {
                        "address": ccRecipientEmailAddress,
                        "name": ccRecipientName
                    }
                }],
                "customizedMessageBody": customizedMessageBody
            }
        };
        return callGraphApi(msalconfig.graphInvitationEndpoint, "POST", body)
                .then(function (response) {
                    //console.log("inviteUser Response: " + response);
                    var contentType = response.headers.get("content-type");
                    if ((response.status === 200 || response.status === 201) && contentType && contentType.indexOf("application/json") !== -1) {
                        return response.json()
                            .then(function (data) {
                                //console.log(data);
                                console.log("Invitation Sent to " + data.invitedUserEmailAddress)
                                return data;
                            })
                            .catch(function (error) {
                                console.log("inviteUser json Error: " + error.message);
                                return error;
                            });
                    } else {
                        return response.json()
                            .then(function (data) {
                                console.log("Invitation Failed: " + data.error.message);
                                return data;
                            })
                            .catch(function (error) {
                                console.log("inviteUser json2 Error: " + error.message);
                                return error;
                            });
                    }
                })
                .catch(function (error) {
                    console.log("inviteUser Error: " + error);
                    return error;
                });
    }
    myGraph.inviteUser = inviteUser;

    /*
    * Add a User to a Group.
    * 
    * @param {any} groupId - GroupId GUID
    * @param {any} userId - UserId GUID
    */
    var addGroupMember = function addGroupMember(groupId, userId) {
        // add user to group
        var endpoint = msalconfig.graphGroupMembersEndpoint.replace("{groupId}", groupId);
        var body = {"@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/" + userId};
        return callGraphApi(endpoint, "POST", body)
                .then(function (response) {
                    var contentType = response.headers.get("content-type");
                    if ((response.status === 204)) {
                        //responseElement.innerHTML += "User Added to Group";
                        console.log("User Added to Group");
                        return Promise.resolve(true);  // no reponse data so return resolved promise
                    } else {
                        return response.json()
                            .then(function (data) {
                                if (response.status === 400 && data.error.message === "One or more added object references already exist for the following modified properties: 'members'.") {
                                    //responseElement.innerHTML += "User Already in Group";
                                    console.log("User Already in Group");
                                    //return data;
                                    return Promise.resolve(true); //returns true if user already in group
                                } else {
                                    console.log("Add Member Failed: " + data.error.message);
                                    return data;
                                }
                            })
                            .catch(function (error) {
                                console.log("Add Member json Failed: " + error.message);
                                return error;
                            });
                    }
                })
                .catch(function (error) {
                    showError(endpoint, error);
                });
    }
    myGraph.addGroupMember = addGroupMember;

     /*
     * Get an access token then Call a Web API .
     * 
     * @param {any} endpoint - Web API endpoint
     * @param {any} method - http method ["GET", "POST"]
     * @param {any} body - http body for POST
     * @param {object} headers - optional - new Headers()
     */
    var callGraphApi = function callGraphApi(endpoint, method, body, headers) {
        return getToken()
        .then(function (token) {
            return new Promise((resolve, reject) => {
                if (!headers) { headers = new Headers(); }
                var bearer = "Bearer " + token;
                headers.append("Authorization", bearer);
                headers.append("Content-Type", 'application/json');
                var options = {
                    method: method,
                    headers: headers,
                };
                if (body) { options.body = JSON.stringify(body); };
                fetch(endpoint, options)
                    .then(function (response) {
                        resolve(response);
                    })
                    .catch(function (error) {
                        reject(error);
                    });
            });
        })
        .catch(function (err) {
            return err;
        });
    }
    myGraph.callGraphApi = callGraphApi;

    /**
    * login and get access token
    */
    var getToken = function getToken() {
        return new Promise((resolve, reject) => {
            var user = userAgentApplication.getUser();
            if (!user) {
                // If user is not signed in, then prompt user to sign in via loginRedirect.
                // This will redirect user to the Azure Active Directory v2 Endpoint
                userAgentApplication.loginRedirect(graphAPIScopes)
            } else {
                //// In order to call the Graph API, an access token needs to be acquired.
                //// Try to acquire the token used to Query Graph API silently first
                userAgentApplication.acquireTokenSilent(graphAPIScopes)
                    .then(function (token) {
                        //After the access token is acquired, return the token
                        resolve(token);
                    }, function (error) {
                        // If the acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
                        // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user 
                        // can re-type the current username and password and/ or give consent to new permissions your application is requesting.
                        // After authentication/ authorization completes, this page will be reloaded again and callGraphApi() will be called.
                        // Then, acquireTokenSilent will then acquire the token silently, the Graph API call results will be made and results will be displayed in the page.
                        if (error) {
                            userAgentApplication.acquireTokenRedirect(graphAPIScopes);
                            reject(error);
                        }
                    });
            }
        });
    }
    myGraph.getToken = getToken;

    /**
    * Show an error message in the page
    * @param {string} endpoint - the endpoint used for the error message
    * @param {string} error - the error string
    * @param {object} errorElement - the HTML element in the page to display the error
    */
    function showError(endpoint, error, errorDesc) {
        var formattedError = JSON.stringify(error, null, 4);
        if (formattedError.length < 3) {
            formattedError = error;
        }
        //document.getElementById("errorMessage").innerHTML = "An error has occurred:<br/>Endpoint: " + endpoint + "<br/>Error: " + formattedError + "<br/>" + errorDesc;
        //console.error(error);
    }

    /**
    * Callback method from sign-in: if no errors, call callGraphApi() to show results.
    * @param {string} errorDesc - If error occur, the error message
    * @param {object} token - The token received from login
    * @param {object} error - The error 
    * @param {string} tokenType - the token type: usually id_token
    */
    function loginCallback(errorDesc, token, error, tokenType) {
        if (errorDesc) {
            showError(msal.authority, error, errorDesc);
        } else {
            callGraphApi();
        }
    }

    var isLoggedIn = function isLoggedIn() {
        var user = userAgentApplication.getUser();
        if (!user) {
            return false;
        } else {
            return true;
        }
    };
    myGraph.isLoggedIn = isLoggedIn;

    /**
    * login the user
    */
    var login = function login() {
        userAgentApplication.loginRedirect(graphAPIScopes);
    }
    myGraph.login = login;

    /**
    * Sign-out the user
    */
    var logout = function logout() {
        userAgentApplication.logout();
    }
    myGraph.logout = logout;

})(this);