<!-- taskpane.js -->
const msalConfig = {
  auth: {
    clientId: "YOUR_CLIENT_ID", // Replace with your Azure AD App's Client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://yourdomain.com/taskpane.html",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
  scopes: ["User.Read"]
};

document.getElementById("signin").onclick = async () => {
  try {
    const result = await msalInstance.loginPopup(loginRequest);
    document.getElementById("output").textContent = `Hello, ${result.account.username}`;
  } catch (error) {
    console.error(error);
    document.getElementById("output").textContent = error.message;
  }
};
