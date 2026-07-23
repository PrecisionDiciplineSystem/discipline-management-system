const msalConfig = {
    auth: {
        clientId: "8942b408-2646-4bbd-bb4f-7a82226351c2",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);


async function signIn() {

    const loginResponse = await msalInstance.loginPopup({
        scopes: [
            "User.Read"
        ]
    });

    console.log(loginResponse.account);

}
