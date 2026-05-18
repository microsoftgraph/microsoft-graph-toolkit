const TRUSTED_PREFIX = 'https://raw.githubusercontent.com/pnp/mgt-samples/main/';

/**
 * Validates that a URL points to a trusted location within the pnp/mgt-samples repository.
 * Uses URL normalization to prevent path traversal attacks (e.g., '../' segments).
 *
 * @param {string} url - The URL to validate
 * @returns {boolean} Whether the URL is trusted
 */
export const isValidManifestUrl = url => {
  if (!url) return false;
  try {
    const normalized = new URL(url).href;
    return normalized.startsWith(TRUSTED_PREFIX) && !normalized.includes('..');
  } catch {
    return false;
  }
};

export { TRUSTED_PREFIX };
