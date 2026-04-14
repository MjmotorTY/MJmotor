const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function initAuth() {
  // 處理登入後的回調
  await msalInstance.handleRedirectPromise();

  const accounts = msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    // 沒有登入紀錄，跳轉登入頁
    msalInstance.loginRedirect({ scopes: CONFIG.scopes });
    return null;
  }

  return accounts[0];
}

async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return null;

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: CONFIG.scopes,
      account: accounts[0],
    });
    return result.accessToken;
  } catch {
    msalInstance.loginRedirect({ scopes: CONFIG.scopes });
    return null;
  }
}

async function getUserInfo() {
  const token = await getToken();
  if (!token) return null;

  const response = await fetch(
    'https://graph.microsoft.com/v1.0/me?$select=displayName,officeLocation,userPrincipalName',
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const user = await response.json();

  return {
    name: user.displayName,           // → 姓名
    unit: user.officeLocation,        // → 所屬據點（中壢所）
    email: user.userPrincipalName,    // → Email
  };
}