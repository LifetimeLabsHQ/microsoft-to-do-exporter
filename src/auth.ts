import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type Configuration,
  type RedirectRequest,
} from '@azure/msal-browser';

export const DEFAULT_CLIENT_ID = '601fc2b3-f7b2-443d-8f97-a5b2082b56e4';
export const GRAPH_SCOPES = ['Tasks.Read'];

const SELECT_ACCOUNT_KEY = 'pn_mstodo_select_account';

let msalInstance: PublicClientApplication | null = null;
let pendingSelectAccount = false;

function getClientId(): string {
  const envId = (import.meta.env.VITE_MS_CLIENT_ID as string | undefined)?.trim();
  return envId || DEFAULT_CLIENT_ID;
}

function readSelectAccountFlag(): boolean {
  if (pendingSelectAccount) return true;
  try {
    return sessionStorage.getItem(SELECT_ACCOUNT_KEY) === '1';
  } catch {
    return false;
  }
}

function writeSelectAccountFlag(value: boolean): void {
  pendingSelectAccount = value;
  try {
    if (value) {
      sessionStorage.setItem(SELECT_ACCOUNT_KEY, '1');
    } else {
      sessionStorage.removeItem(SELECT_ACCOUNT_KEY);
    }
  } catch {
    // Ignore storage restriction errors.
  }
}

export async function initAuth(): Promise<AccountInfo | null> {
  if (!msalInstance) {
    const config: Configuration = {
      auth: {
        clientId: getClientId(),
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: window.location.origin,
        postLogoutRedirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
      },
    };
    msalInstance = new PublicClientApplication(config);
    await msalInstance.initialize();
  }

  const redirectResult = await msalInstance.handleRedirectPromise();
  if (redirectResult?.account) {
    writeSelectAccountFlag(false);
    msalInstance.setActiveAccount(redirectResult.account);
    return redirectResult.account;
  }

  const existing = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0] ?? null;
  if (existing) {
    msalInstance.setActiveAccount(existing);
  }
  return existing;
}

export async function signInWithMicrosoft(forceAccountPicker = false): Promise<void> {
  if (!msalInstance) await initAuth();
  const shouldSelectAccount = forceAccountPicker || readSelectAccountFlag();
  writeSelectAccountFlag(false);
  const request: RedirectRequest = shouldSelectAccount
    ? { scopes: GRAPH_SCOPES, prompt: 'select_account' }
    : { scopes: GRAPH_SCOPES };
  await msalInstance!.loginRedirect(request);
}

export interface AcquireTokenOptions {
  forceRefresh?: boolean;
  allowRedirect?: boolean;
}

/**
 * Returns a Microsoft Graph access token for the signed-in account.
 * Returns `null` when `allowRedirect` triggers an interactive redirect.
 */
export async function acquireGraphToken(options: AcquireTokenOptions = {}): Promise<string | null> {
  const { forceRefresh = false, allowRedirect = true } = options;
  if (!msalInstance) await initAuth();
  const account = msalInstance!.getActiveAccount() ?? msalInstance!.getAllAccounts()[0];
  if (!account) {
    throw new Error('Not signed in. Click "Sign in with Microsoft" first.');
  }
  try {
    const res = await msalInstance!.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account,
      forceRefresh,
    });
    return res.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError && allowRedirect) {
      await msalInstance!.acquireTokenRedirect({
        scopes: GRAPH_SCOPES,
        account,
      });
      return null;
    }
    throw err;
  }
}

export async function signOutLocal(): Promise<void> {
  if (msalInstance) {
    const account = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0];
    msalInstance.setActiveAccount(null);
    if (account) {
      try {
        await msalInstance.clearCache({ account });
      } catch {
        // Cache also cleared via sessionStorage below.
      }
    }
  }
  try {
    sessionStorage.clear();
  } catch {
    // Ignore storage restriction errors.
  }
  // Ensure the next sign-in opens Microsoft's account picker so users can switch accounts.
  writeSelectAccountFlag(true);
}

/** Test-only helper to reset module state between unit tests. */
export function _resetAuthForTests(customInstance?: PublicClientApplication | null): void {
  msalInstance = customInstance ?? null;
  pendingSelectAccount = false;
}
