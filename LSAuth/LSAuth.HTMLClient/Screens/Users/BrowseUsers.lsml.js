/// <reference path="~/GeneratedArtifacts/viewModel.js" />
/// <reference path="../../scripts/jab-msgraphapi-1.0.0.js" />

myapp.BrowseUsers.Invite_execute = function (screen) {
    // Write code here.
    //myGraph.testCall()
    //.then(function (result) {
    //    if (result === true) {
    //        alert("Success!")
    //    } else {
    //        alert("Failure" + result.toString())
    //    }
    //});
    var inviteRedeemUrl = "";

    //this will invite an external user then add the user as a member of the externals group
    myGraph.inviteUser("JoshTest",
                            "jbooker@midcoast.com",
                            msalconfig.redirectUri,
                            true,
                            "Hi Josh2, you can find more information here: " + msalconfig.redirectUri,
                            "josh@joshbooker.com",
                            "ccJosh")
        .then(function (data) {
            //we are now inside the inviteUser callback function
            //inviteUser can completed and the response data is available here
            inviteRedeemUrl = data.inviteRedeemUrl;
            //log the inviteRedeemUrl
            console.log("RedeemUrl: " + inviteRedeemUrl);
            //we're about to call addGroupMember using the response data from inviteUser
            myGraph.addGroupMember(msalconfig.groupId, data.invitedUser.id)
                //addGroupMember returns true on success and error response object on failure
                .then(function (result) {
                    //
                    if (result === true) {
                        alert("addGroupMember Success!")
                    } else {
                        alert("addGroupMember Failed: " + result.error.message)
                    }
                });
        });
};
myapp.BrowseUsers.Invite_canExecute = function (screen) {
    //this hides the Invite button when the user is not logged in
    //it's good practice to hide any button that calls GraphApi 
    // Invite can execute if user isLoggedIn returns true
    return myGraph.isLoggedIn();
};
myapp.BrowseUsers.Login_canExecute = function (screen) {
    //this hides the Login button when the use is already logged in
    // Login can execute when user isLoggedIn returns false
    return !myGraph.isLoggedIn();
};
myapp.BrowseUsers.Login_execute = function (screen) {
    //this logs in the user
    //this is descructive as the page will redirect to the home page of this app
    //we could use query strings to save state, etc
    myGraph.login();
};
myapp.BrowseUsers.Logout_canExecute = function (screen) {
    //this hides the Logout button when the use not yet logged in
    // Logout can execute when user isLoggedIn returns true
    return myGraph.isLoggedIn();
};
myapp.BrowseUsers.Logout_execute = function (screen) {
    //this logs in the user
    //this is descructive as the page will redirect to logout page
    //i beleive there is a setting to get it to redirect back to app afer logout but....(?)
    myGraph.logout();
};