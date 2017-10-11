/// <reference path="adal.js" />
/// <reference path="msal.js" />
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
    var myGraph = window.myGraph || {};

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
    //ea6cd5ab-9c40-4e27-b86b-8aaf86bae36a
    //https://login.microsoftonline.com/common/adminconsent?client_id=39683401-a60c-4bc6-b744-6dbaedef498d&state=12345&redirect_uri=http://localhost:57245/
    //https://login.microsoftonline.com/common/adminconsent?client_id=ea6cd5ab-9c40-4e27-b86b-8aaf86bae36a&state=12345&redirect_uri=http://localhost:57245/

// Graph API scope used to obtain the access token to read user profile
var graphAPIScopes = ["https://graph.microsoft.com/user.read"];
myGraph.graphAPIScopes = graphAPIScopes;

// Initialize application
var userAgentApplication = new Msal.UserAgentApplication(msalconfig.clientID, msalconfig.authorityHostUrl, loginCallback, {
    redirectUri: msalconfig.redirectUri
});

var authContext = new AuthenticationContext(adalconfig);
window.authContext = authContext;

//var token = authContext.acquireToken(adalconfig.clientId, function (a, b, c) { return [a, b, c]; });
 
    //alert(token());

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
POST https://graph.microsoft.com/beta/me/sendMail
Content-type: application/json
Content-length: 512

{
  "message": {
    "subject": "Meet for lunch?",
    "body": {
      "contentType": "Text",
      "content": "The new cafeteria is open."
    },
    "toRecipients": [
      {
        "emailAddress": {
          "address": "samanthab@contoso.onmicrosoft.com"
        }
      }
    ],
    "ccRecipients": [
      {
        "emailAddress": {
          "address": "danas@contoso.onmicrosoft.com"
        }
      }
    ]
  },
  "saveToSentItems": "false"
}

*/


    /*
    * Send Email on behalf or current user
    * 
    * @param {any} subject - subject
    * @param {any} bodyContentType - "text" or "HTML"
    * @param {any} bodyContent - bodyContent
    * @param {any} toRecipientEmail - TO email address
    * @param {any} ccRecipientEmail - CC email address
    * @param {any} saveToSentItems - true or false
    */
var sendMail = function sendMail(subject, bodyContentType, bodyContent, toRecipientEmail, ccRecipientEmail, saveToSentItems) {
    var body = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": bodyContentType,
                "content": bodyContent
            },
            "toRecipients": [
              {
                  "emailAddress": {
                      "address": toRecipientEmail
                  }
              }
            ],
            "ccRecipients": [
              {
                  "emailAddress": {
                      "address": ccRecipientEmail
                  }
              }
            ]
        },
        "saveToSentItems": saveToSentItems
    };
    return callGraphApi(msalconfig.graphUsersEndpoint + "/1bb51a8d-9472-43e0-8166-213de80817da/sendMail", "POST", body)
            .then(function (response) {
                var contentType = response.headers.get("content-type");
                if ((response.status === 202)) {
                    console.log("Email Sent Successfully");
                    return Promise.resolve(true);  // no reponse data so return resolved promise
                } else {
                    return response.json()
                        .then(function (data) {
                            if (response.status === 400 && data.error.message === "<possible sendMail specific error code to handle here>") {
                                console.log("some email error message");
                                //return data;
                                return Promise.resolve(true); //returns true if user already in group
                            } else {
                                console.log("Send Mail Failed: " + data.error.message);
                                return data;
                            }
                        })
                        .catch(function (error) {
                            console.log("Send Mail json Failed: " + error.message);
                            return error;
                        });
                }
            })
            .catch(function (error) {
                console.log("sendMail Error: " + error.message);
                return error;
            });
}
myGraph.sendMail = sendMail;

    /*
* Send Email on behalf or current user
* 
* @param {any} subject - subject
* @param {any} bodyContentType - "text" or "HTML"
* @param {any} bodyContent - bodyContent
* @param {any} toRecipientEmail - TO email address
* @param {any} ccRecipientEmail - CC email address
* @param {any} saveToSentItems - true or false
*/
var sendMailServer = function sendMailServer(subject, bodyContentType, bodyContent, toRecipientEmail, ccRecipientEmail, saveToSentItems, fromUserId) {
    var body = {
        "method": "sendMail",
        "subject": subject,
        "bodyContentType": bodyContentType,
        "bodyContent": bodyContent,
        "toRecipientEmail": toRecipientEmail,
        "ccRecipientEmail": ccRecipientEmail,
        "saveToSentItems": saveToSentItems,
        "fromUserId": fromUserId
    };
    return callGraphClient(body)
            .then(function (response) {
                if ((response.status === 202)) {
                    console.log("Email Sent Successfully");
                    return Promise.resolve(true);  // no reponse data so return resolved promise
                } else {
                    return response.json()
                        .then(function (data) {
                                console.log("Send Mail Failed: " + data.error.message);
                                return data;
                        })
                        .catch(function (error) {
                            console.log("Send Mail json Failed: " + error.message);
                            return error;
                        });
                }
            })
            .catch(function (error) {
                console.log("sendMail Error: " + error.message);
                return error;
            });
}
myGraph.sendMailServer = sendMailServer;

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
     * Get app access token then Call a GraphApi.
     * 
     * @param {any} endpoint - Web API endpoint
     * @param {any} method - http method ["GET", "POST"]
     * @param {any} body - http body for POST
     * @param {object} headers - optional - new Headers()
     */
    var callGraphApi = function callGraphApi(endpoint, method, body, headers) {
        return getAppToken()
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
    * sends request to Node.js Azure Function which will aquire an id_Token and then call GraphApi and return the response
    * @param {any} body - JSON body with parameters.  "method" : "getToken" or "sendMail"
    *
    * example body for getToken:
    *   { "method": "getToken" }
    *
    * example body for sendMail:
    *   {
    *        "method": "sendMail",
    *        "subject": "Meet for lunch?",
    *        "bodyContentType": "Text",
    *        "bodyContent": "The new cafeteria is open.",
    *        "toRecipientEmail": "jbooker@midcoast.com",
    *        "ccRecipientEmail": "josh@joshuabooker.com",
    *        "saveToSentItems": "true",
    *        "fromUserId": "amanda@joshbooker.com"
    *   }
    *
    */
    var callGraphClient = function callGraphClient(body) {
        return new Promise((resolve, reject) => {
            var options = {
                url: msalconfig.graphClientEndpoint,
                headers: { "Content-Type": "application/json" },
                method: "POST",
                data: JSON.stringify(body)
            };
            $.ajax(options)
                .then(function (Response) {
                    resolve(Response);
                });
        });
    }
    myGraph.callGraphClient = callGraphClient;


    /**
    * login and get user access token
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
    * sends request to Node.js Azure Function to get id_token via adal-node method: adal.AuthenticationContext.acquireTokenWithClientCredentials()
    */
    var getAppToken = function getAppToken() {
        return new Promise((resolve, reject) => {
            var endpoint = msalconfig.graphClientEndpoint;
            var options = {
                url: endpoint,
                headers: { "Content-Type": "application/json" },
                method: "POST",
                data: JSON.stringify({ method: "getToken" })
            };
            $.ajax(options)
                .then(function (tokenRes) {
                    var accesstoken = tokenRes.accessToken;
                    resolve(accesstoken);
                });
        });
    }
    myGraph.getAppToken = getAppToken;

    var getAppToken2 = function getAppToken2() {

        return new Promise((resolve, reject) => {
            var endpoint = msalconfig.graphClientEndpoint
            //var endpoint = msalconfig.graphClientEndpoint;
            //var method = "GET"
            //var headers = new Headers();
            //headers.append("Content-Type", 'text/plain');
            //var options = {
            //    headers: headers,
            //    method: method
            //    //,mode: 'cors'
            //};
            //fetch(endpoint, options)
            //    .then(function (tokenRes) {
            //        var accesstoken = tokenRes.accessToken;
            //        resolve(accesstoken);
            //    })
            //    .catch(function (error) {
            //        reject(error);
            //    });
            // cannot get fetch to work with cors - lets try ajax
            var options = {
                url: endpoint,
                headers: {"Content-Type": "application/json"},
                method: "POST",
                data: JSON.stringify({ method: "getToken" })
            };
            $.ajax(options)
                .then(function (tokenRes) {
                    var accesstoken = tokenRes.accessToken;
                    resolve(accesstoken);
                });
        });
    }
    myGraph.getAppToken2 = getAppToken2;

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