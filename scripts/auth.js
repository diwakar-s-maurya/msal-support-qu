const CLIENT_ID_WORKING = "d63e5128-a9f4-437c-9d8f-afefb9a9658f"
const CLIENT_ID_NOT_WORKING = "3bfed8ab-433b-4014-902c-8a6932ea8bde"

const msalParams = {
    auth: {
        authority:  "https://login.microsoftonline.com/eb1f94dc-8d25-4c67-985c-04be74c8f698",
        clientId: CLIENT_ID_NOT_WORKING, // Switching this to CLIENT_ID_WORKING, makes everything work
        redirectUri: "http://localhost:3000",
    },
}

function combine(...paths) {
    return paths
        .map(path => path.replace(/^[\\|/]/, "").replace(/[\\|/]$/, ""))
        .join("/")
        .replace(/\\/g, "/");
}

async function getToken() {

    const command = {
        resource: "https://betagro-my.sharepoint.com/",
        command: "authenticate",
        type: "SharePoint",
    }

    const app = new msal.PublicClientApplication(msalParams);

    let accessToken = "";
    let authParams = { scopes: [] };

    switch (command.type) {
        case "SharePoint":
        case "SharePoint_SelfIssued":
            authParams = { scopes: [`${combine(command.resource, ".default")}`] };
            break;
        default:
            break;
    }

    console.log("Calling with these auth params", authParams);

    try {
        const resp = await app.acquireTokenSilent(authParams);
        accessToken = resp.accessToken;

    } catch (e) {
        const resp = await app.loginPopup(authParams);
        app.setActiveAccount(resp.account);

        if (resp.idToken) {
            const resp2 = await app.acquireTokenSilent(authParams);
            accessToken = resp2.accessToken;
        }
    }

    console.log("token", accessToken);
    return accessToken;
}
