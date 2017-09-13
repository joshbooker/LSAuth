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

    myGraph.inviteUser("JoshTest",
                            "jbooker@midcoast.com",
                            msalconfig.redirectUri,
                            true,
                            "Hi Josh2, you can find more information here: " + msalconfig.redirectUri,
                            "josh@joshbooker.com",
                            "ccJosh")
        .then(function (data) {
            //invite response data is available here
            inviteRedeemUrl = data.inviteRedeemUrl;
            //log the inviteRedeemUrl
            console.log("RedeemUrl: " + inviteRedeemUrl);
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
    // Write code here.
    return myGraph.isLoggedIn();
};
myapp.BrowseUsers.Login_canExecute = function (screen) {
    // Write code here.
    return !myGraph.isLoggedIn();
};
myapp.BrowseUsers.Login_execute = function (screen) {
    // Write code here.
    myGraph.login();
};
myapp.BrowseUsers.Logout_canExecute = function (screen) {
    // Write code here.
    return myGraph.isLoggedIn();
};
myapp.BrowseUsers.Logout_execute = function (screen) {
    // Write code here.
    myGraph.logout();
};