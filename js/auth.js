// MSAL 認証モジュール
let msalInstance = null;

async function initAuth() {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();

    // ポップアップからの戻りを処理
    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            msalInstance.setActiveAccount(response.account);
        }
    } catch (error) {
        console.error("handleRedirectPromise error:", error);
    }
}

function getActiveAccount() {
    const active = msalInstance.getActiveAccount();
    if (active) return active;

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        return accounts[0];
    }
    return null;
}

async function signIn() {
    try {
        const response = await msalInstance.loginPopup({
            scopes: graphConfig.scopes,
            prompt: "select_account",
        });
        msalInstance.setActiveAccount(response.account);
        return response.account;
    } catch (error) {
        if (error.errorCode === "popup_window_error" || error.errorCode === "empty_window_error") {
            throw new Error("popup_blocked");
        }
        if (error.errorCode === "user_cancelled") {
            throw new Error("user_cancelled");
        }
        console.error("Login failed:", error);
        throw error;
    }
}

async function getAccessToken() {
    const account = getActiveAccount();
    if (!account) throw new Error("No active account");

    const request = {
        scopes: graphConfig.scopes,
        account: account,
    };

    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            const response = await msalInstance.acquireTokenPopup(request);
            return response.accessToken;
        }
        throw error;
    }
}

// サイレント専用トークン取得（ポップアップを開かない）
// 権限がなければ null を返す
async function getAccessTokenSilentOnly(scopes) {
    const account = getActiveAccount();
    if (!account) return null;
    try {
        const response = await msalInstance.acquireTokenSilent({ scopes, account });
        return response.accessToken;
    } catch {
        return null;
    }
}

async function signOut() {
    const account = getActiveAccount();
    if (account) {
        await msalInstance.logoutPopup({
            account: account,
        });
    }
}
