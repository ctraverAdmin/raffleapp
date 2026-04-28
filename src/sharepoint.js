import { PublicClientApplication } from "@azure/msal-browser";

const CLIENT_ID = "b3c0ddf3-9dce-4cbb-bbd0-35510919ab94";
const TENANT_ID = "c3a38b12-13d7-4b38-9c9f-f6cad2952b44";

const SITE_HOSTNAME = "northeastdata1.sharepoint.com";
const SITE_PATH = "/sites/TCAA";

export const SUMMARY_LIST_NAME = "Raffle Summary List";
export const RECEIPTS_LIST_NAME = "Raffle Receipts";

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
};

const msalApp = new PublicClientApplication(msalConfig);
let initialized = false;

async function initMsal() {
  if (!initialized) {
    await msalApp.initialize();

    const response = await msalApp.handleRedirectPromise();

    if (response?.account) {
      msalApp.setActiveAccount(response.account);
    } else {
      const accounts = msalApp.getAllAccounts();
      if (accounts.length > 0) {
        msalApp.setActiveAccount(accounts[0]);
      }
    }

    initialized = true;
  }
}

export async function signIn() {
  await initMsal();

  const account = msalApp.getActiveAccount();

  if (account) {
    return account;
  }

  await msalApp.loginRedirect({
    scopes: ["Sites.ReadWrite.All"],
  });
}

export async function getToken() {
  await initMsal();

  let account = msalApp.getActiveAccount();

  if (!account) {
    await signIn();
    return null;
  }

  try {
    const token = await msalApp.acquireTokenSilent({
      account,
      scopes: ["Sites.ReadWrite.All"],
    });

    return token.accessToken;
  } catch {
    await msalApp.acquireTokenRedirect({
      scopes: ["Sites.ReadWrite.All"],
    });

    return null;
  }
}

async function graphFetch(url, options = {}) {
  const token = await getToken();

  if (!token) {
    throw new Error("Login is still processing. Click Connect to SharePoint again.");
  }

  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || response.statusText);
  }

  if (response.status === 204) return null;
  return response.json();
}

export async function getSiteId() {
  const site = await graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_HOSTNAME}:${SITE_PATH}`
  );

  return site.id;
}

export async function getListId(siteId, listName) {
  const lists = await graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName`
  );

  const list = lists.value.find((item) => item.displayName === listName);

  if (!list) {
    throw new Error(`Could not find SharePoint list: ${listName}`);
  }

  return list.id;
}

export async function getColumnMap(siteId, listId) {
  const columns = await graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`
  );

  const map = {};

  columns.value.forEach((col) => {
    map[col.displayName] = col.name;
  });

  return map;
}

function makeFields(columnMap, values) {
  const fields = {
    Title: values.Title || "",
  };

  Object.entries(values).forEach(([displayName, value]) => {
    if (displayName === "Title") return;

    const internalName = columnMap[displayName];

    if (internalName) {
      fields[internalName] = value;
    }
  });

  return fields;
}

export async function getListItems(siteId, listId) {
  const result = await graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&top=999`
  );

  return result.value || [];
}

export async function clearList(siteId, listId) {
  const items = await getListItems(siteId, listId);

  for (const item of items) {
    await graphFetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${item.id}`,
      {
        method: "DELETE",
      }
    );
  }
}

export async function createListItem(siteId, listId, columnMap, values) {
  return graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
    {
      method: "POST",
      body: JSON.stringify({
        fields: makeFields(columnMap, values),
      }),
    }
  );
}

export async function connectSharePoint() {
  const siteId = await getSiteId();
  const summaryListId = await getListId(siteId, SUMMARY_LIST_NAME);
  const receiptsListId = await getListId(siteId, RECEIPTS_LIST_NAME);

  const summaryColumnMap = await getColumnMap(siteId, summaryListId);
  const receiptsColumnMap = await getColumnMap(siteId, receiptsListId);

  return {
    siteId,
    summaryListId,
    receiptsListId,
    summaryColumnMap,
    receiptsColumnMap,
  };
}