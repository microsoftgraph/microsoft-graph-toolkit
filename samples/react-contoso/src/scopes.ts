/**
 * Object mapping chat operations to the scopes required to perform them
 */
const dashboardOperationScopes: Record<string, string[]> = {
  tasks: ['tasks.readwrite']
};

/**
 * Provides an array of the distinct scopes required for all chat operations
 */
export const allDashboardScopes = Array.from(
  Object.values(dashboardOperationScopes).reduce((acc, scopes) => {
    scopes.forEach(s => acc.add(s));
    return acc;
  }, new Set<string>())
);
