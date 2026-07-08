/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Signed-in account identity (U5 / D7). Durable state — approval tokens (D8) and
 * the durable-ID alias table (D3) — is scoped per account so a token minted
 * while signed in as one user can never be redeemed as another. MSAL's
 * `homeAccountId` (`<oid>.<tid>`) is the stable per-account key.
 *
 * `getAccount()` is async (it reads the MSAL cache), but the token manager and
 * resolver need a scope key synchronously at call time. So the resolved id is
 * memoized after sign-in and read back via {@link currentAccountId}. Only a real
 * id is memoized — an unauthenticated lookup returns the `'default'` fallback
 * WITHOUT caching, so the real id is still picked up once sign-in completes.
 */

import { getAccount } from './device-code-flow.js';

/** Scope key used before sign-in (or when the account can't be determined). */
export const DEFAULT_ACCOUNT_ID = 'default';

let cachedAccountId: string | null = null;

/**
 * Resolves the signed-in account's stable id from MSAL, memoizing the first real
 * value. Returns {@link DEFAULT_ACCOUNT_ID} (uncached) when unauthenticated.
 */
export async function resolveAccountId(): Promise<string> {
  if (cachedAccountId !== null) {
    return cachedAccountId;
  }
  const account = await getAccount();
  const homeAccountId = account?.homeAccountId;
  if (homeAccountId != null && homeAccountId.length > 0) {
    cachedAccountId = homeAccountId;
    return cachedAccountId;
  }
  return DEFAULT_ACCOUNT_ID;
}

/**
 * The last resolved account id without touching the MSAL cache. Returns
 * {@link DEFAULT_ACCOUNT_ID} until {@link resolveAccountId} has populated it.
 */
export function currentAccountId(): string {
  return cachedAccountId ?? DEFAULT_ACCOUNT_ID;
}

/** Clears the memo so a sign-out / account switch re-resolves the identity. */
export function resetAccountId(): void {
  cachedAccountId = null;
}
