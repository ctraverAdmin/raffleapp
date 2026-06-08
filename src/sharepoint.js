import { PublicClientApplication } from "@azure/msal-browser";

const CLIENT_ID = "b3c0ddf3-9dce-4cbb-bbd0-35510919ab94";
const TENANT_ID = "c3a38b12-13d7-4b38-9c9f-f6cad2952b44";

const SITE_HOSTNAME = "northeastdata1.sharepoint.com";
const SITE_PATH = "/sites/TCAA";

export const SUMMARY_LIST_NAME = "Raffle Summary List";
export const RECEIPTS_LIST_NAME = "Raffle Receipts";
export const CASH_SALES_LIST_NAME = "Raffle Cash Sales";
export const SPECIAL_FUNDRAISERS_LIST_NAME = "Special Fundraisers";
export const TREASURER_TRANSFERS_LIST_NAME = "Treasurer Transfers";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

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
  if (initialized) return;

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

export async function signIn() {
  await initMsal();

  const account = msalApp.getActiveAccount();

  if (account) {
    return account;
  }

  await msalApp.loginRedirect({
    scopes: ["Sites.ReadWrite.All"],
  });

  return null;
}

export async function getToken() {
  await initMsal();

  const account = msalApp.getActiveAccount();

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
  } catch (error) {
    console.error("Silent token failed:", error);

    await msalApp.acquireTokenRedirect({
      scopes: ["Sites.ReadWrite.All"],
    });

    return null;
  }
}

function getGraphErrorMessage(errorText, fallbackMessage) {
  if (!errorText) return fallbackMessage;

  try {
    const parsed = JSON.parse(errorText);
    return (
      parsed?.error?.message ||
      parsed?.error_description ||
      errorText ||
      fallbackMessage
    );
  } catch {
    return errorText || fallbackMessage;
  }
}

function isItemNotFoundError(status, errorText) {
  if (status === 404) return true;

  if (!errorText) return false;

  try {
    const parsed = JSON.parse(errorText);
    const code = String(parsed?.error?.code || "").toLowerCase();
    const message = String(parsed?.error?.message || "").toLowerCase();

    return (
      code === "itemnotfound" ||
      message.includes("specified list item was not found") ||
      message.includes("list item was not found") ||
      message.includes("item not found")
    );
  } catch {
    const lower = String(errorText).toLowerCase();

    return (
      lower.includes("itemnotfound") ||
      lower.includes("specified list item was not found") ||
      lower.includes("list item was not found") ||
      lower.includes("item not found")
    );
  }
}

async function graphFetch(url, options = {}) {
  const token = await getToken();

  if (!token) {
    throw new Error(
      "Login is still processing. Click Connect to SharePoint again."
    );
  }

  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (response.status === 204) {
    return null;
  }

  const text = await response.text();

  if (!response.ok) {
    const message = getGraphErrorMessage(text, response.statusText);

    throw new Error(message);
  }

  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

async function graphDelete(url) {
  const token = await getToken();

  if (!token) {
    throw new Error(
      "Login is still processing. Click Connect to SharePoint again."
    );
  }

  const response = await fetch(url, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (response.ok || response.status === 204) {
    return;
  }

  const errorText = await response.text();

  if (isItemNotFoundError(response.status, errorText)) {
    console.warn("SharePoint item already deleted. Continuing save.", {
      url,
      status: response.status,
      errorText,
    });
    return;
  }

  const message = getGraphErrorMessage(
    errorText,
    `Could not delete SharePoint item. Status ${response.status}.`
  );

  throw new Error(message);
}

export async function getSiteId() {
  const site = await graphFetch(
    `${GRAPH_BASE}/sites/${SITE_HOSTNAME}:${SITE_PATH}`
  );

  return site.id;
}

export async function getListId(siteId, listName) {
  const lists = await graphFetch(
    `${GRAPH_BASE}/sites/${siteId}/lists?$select=id,displayName`
  );

  const list = (lists.value || []).find(
    (item) => item.displayName === listName
  );

  if (!list) {
    throw new Error(`Could not find SharePoint list: ${listName}`);
  }

  return list.id;
}

export async function getColumnMap(siteId, listId) {
  const columns = await graphFetch(
    `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`
  );

  const map = {};

  (columns.value || []).forEach((col) => {
    map[col.displayName] = col.name;
  });

  return map;
}

function makeFields(columnMap, values) {
  const fields = {
    Title: values.Title || "",
  };

  Object.entries(values || {}).forEach(([displayName, value]) => {
    if (displayName === "Title") return;

    const internalName = columnMap?.[displayName];

    if (internalName) {
      fields[internalName] = value;
    } else {
      console.warn(
        `SharePoint column was not found and will be skipped: ${displayName}`
      );
    }
  });

  return fields;
}

export async function getListItems(siteId, listId) {
  let allItems = [];

  let url =
    `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items` +
    `?$expand=fields&$top=999`;

  while (url) {
    const result = await graphFetch(url);

    allItems = [...allItems, ...(result?.value || [])];

    url = result?.["@odata.nextLink"] || null;
  }

  return allItems;
}

export async function clearList(siteId, listId) {
  let url =
    `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items` +
    `?$select=id&$top=999`;

  while (url) {
    const result = await graphFetch(url);
    const items = result?.value || [];

    for (const item of items) {
      if (!item?.id) continue;

      await graphDelete(
        `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items/${item.id}`
      );
    }

    url = result?.["@odata.nextLink"] || null;
  }
}

export async function createListItem(siteId, listId, columnMap, values) {
  const fields = makeFields(columnMap, values);

  return graphFetch(`${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    body: JSON.stringify({
      fields,
    }),
  });
}

export async function connectSharePoint() {
  const siteId = await getSiteId();

  const summaryListId = await getListId(siteId, SUMMARY_LIST_NAME);
  const receiptsListId = await getListId(siteId, RECEIPTS_LIST_NAME);
  const cashSalesListId = await getListId(siteId, CASH_SALES_LIST_NAME);

  const specialFundraisersListId = await getListId(
    siteId,
    SPECIAL_FUNDRAISERS_LIST_NAME
  );

  const transfersListId = await getListId(
    siteId,
    TREASURER_TRANSFERS_LIST_NAME
  );

  const summaryColumnMap = await getColumnMap(siteId, summaryListId);
  const receiptsColumnMap = await getColumnMap(siteId, receiptsListId);
  const cashSalesColumnMap = await getColumnMap(siteId, cashSalesListId);

  const specialFundraisersColumnMap = await getColumnMap(
    siteId,
    specialFundraisersListId
  );

  const transfersColumnMap = await getColumnMap(siteId, transfersListId);

  return {
    siteId,

    summaryListId,
    receiptsListId,
    cashSalesListId,
    specialFundraisersListId,
    transfersListId,

    summaryColumnMap,
    receiptsColumnMap,
    cashSalesColumnMap,
    specialFundraisersColumnMap,
    transfersColumnMap,
  };
}