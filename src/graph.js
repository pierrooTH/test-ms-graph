import { graphConfig } from "./authConfig";

/**
 * Attaches a given access token to a MS Graph API call. Returns information about the user
 * @param accessToken 
 */
export async function callMsGraph(accessToken) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(graphConfig.graphMeEndpoint, options)
        .then(response => response.json())
        .catch(error => console.log(error));
}

export async function callMsGraphEmails(accessToken) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(graphConfig.graphMeMailFolder, options)
        .then(async response => {
            const res = await response.json();
            const id = res.value[0].id;
            return fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${id}/messages?$filter=isRead eq false&$count=true`, options)
                .then(response2 => response2.json())
        } )
        .catch(error => console.log(error));
}