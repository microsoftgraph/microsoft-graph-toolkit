# Microsoft Graph Toolkit Development Instructions

Always reference these instructions first and fallback to search or bash commands only when you encounter unexpected information that does not match the info here.

## Working Effectively

### Initial Setup (Required for fresh repository clone)
- `PUPPETEER_SKIP_DOWNLOAD=true yarn install --immutable` -- installs all dependencies. Takes ~30 seconds. Skip puppeteer due to common network issues.
- `yarn playwright install chromium --with-deps` -- installs browser dependencies for testing. Takes ~5-10 minutes. NEVER CANCEL. Set timeout to 15+ minutes.
  - NOTE: May fail due to network/firewall limitations. Tests require this, but build/development work without it.
- `yarn build` -- full build of all packages. Takes ~60 seconds. NEVER CANCEL. Set timeout to 90+ minutes.

### Core Development Commands
- `yarn build` -- builds all packages with prettier check and cleaning. Takes ~60 seconds. NEVER CANCEL. Set timeout to 90+ seconds.
- `yarn build:dev` -- faster build excluding some packages (mgt, spfx, sharepoint, electron, teamsfx, proxy providers). Takes ~45 seconds.
- `yarn build:compile` -- compile-only build (no bundle). Takes ~45 seconds. Used by storybook.
- `yarn clean` -- cleans all dist folders and build artifacts. Takes ~5 seconds.
- `yarn watch` -- starts TypeScript and Sass compilation in watch mode for development. Takes ~10 seconds to start. Runs continuously.
- `yarn lint` -- runs ESLint and Stylelint. Takes ~25 seconds. NEVER CANCEL. Set timeout to 45+ seconds.
- `yarn lint:fix` -- automatically fixes linting issues. Takes ~30 seconds.

### Testing Commands  
- `yarn test` -- runs all tests with coverage using Web Test Runner + Playwright. Takes ~2-5 minutes. NEVER CANCEL. Set timeout to 10+ minutes.
  - REQUIRES: `yarn playwright install chromium --with-deps` to be run first
  - Tests may fail in environments with network restrictions preventing browser downloads
- `yarn wtr` -- runs tests without coverage. Faster than `yarn test`.
- `yarn wtr:watch` -- runs tests in watch mode for development.

### Storybook Development
- `yarn storybook:dev` -- builds components and generates custom-elements.json for storybook. Takes ~45 seconds.
- `yarn storybook` -- starts storybook development server on http://localhost:6006. Takes ~15 seconds to start.
- `yarn storybook:build` -- builds storybook for production. Takes ~2-3 minutes. NEVER CANCEL. Set timeout to 5+ minutes.
- `yarn storybook:bundle` -- bundles storybook components with rollup.
- `yarn storybook:watch` -- runs storybook with watch mode (components + storybook in parallel).

### Development Servers
- `yarn serve` -- starts web dev server on http://localhost:3000 for testing components. Instant startup.
- `yarn serve:https` -- same as serve but with HTTPS on port 3000.
- `yarn start` -- builds and starts development with watch + serve in parallel.
- `yarn watch:serve` -- runs watch and serve in parallel for continuous development.

### Sample Applications
- `yarn watch:react-contoso` -- starts React sample app on http://localhost:3000. Takes ~5 seconds to start.
- `yarn watch:react` -- builds React library dependencies and starts React sample in watch mode.

## Validation Requirements

### ALWAYS test core functionality after changes:
1. **Build Process**: Run `yarn build` and ensure all packages compile successfully
2. **Linting**: Run `yarn lint` and fix any issues with `yarn lint:fix`
3. **Component Testing**: Start storybook with `yarn storybook` and verify components render
4. **Development Server**: Run `yarn serve` and test components in browser at http://localhost:3000
5. **React Integration**: Test React components with `yarn watch:react-contoso`

### Manual Testing Scenarios
After making changes, ALWAYS validate with these scenarios:
- Build and start storybook to verify component rendering: `yarn storybook:dev && yarn storybook`
- Test development server functionality: `yarn serve` and open http://localhost:3000
- Verify React integration works: `yarn watch:react-contoso` and test at http://localhost:3000
- Run linting to ensure code style: `yarn lint`

## Repository Structure

### Key Packages (monorepo with yarn workspaces + lerna)
- `packages/mgt-element/` -- Base classes and core functionality for all components
- `packages/mgt-components/` -- All UI components (mgt-person, mgt-login, mgt-agenda, etc.)
- `packages/mgt-react/` -- React wrapper components (auto-generated from mgt-components)
- `packages/mgt/` -- Bundle package that includes all components
- `packages/providers/` -- Authentication providers (msal2, sharepoint, teamsfx, proxy, electron)
- `samples/react-contoso/` -- React sample application for testing

### Important Files
- `package.json` -- Root package with all scripts and dependencies
- `lerna.json` -- Monorepo configuration for managing packages
- `web-test-runner.config.mjs` -- Test configuration using Playwright
- `.storybook/` -- Storybook configuration and addons
- `tsconfig.base.json` -- TypeScript configuration for all packages

### Build Process Details
1. **Sass Compilation**: Converts .scss files to -css.ts files using `compileSass.js`
2. **TypeScript Compilation**: Builds ES6 modules in each package's dist/ folder
3. **Bundling**: Rollup creates final bundles for distribution
4. **Custom Elements Manifest**: Generates metadata for web components

## Common Issues and Workarounds

### Dependencies
- **Puppeteer Download Issues**: Always use `PUPPETEER_SKIP_DOWNLOAD=true yarn install` for installation
- **Browser Download Failures**: May occur due to firewall/network. Development works without browsers, only testing is affected
- **Peer Dependency Warnings**: Expected and safe to ignore during installation

### Timing Expectations (NEVER CANCEL these operations)
- **Initial Install**: 30-60 seconds (add timeout buffer to 2+ minutes)  
- **Full Build**: 60 seconds (add timeout buffer to 2+ minutes)
- **Linting**: 25 seconds (add timeout buffer to 1+ minute)
- **Tests**: 2-5 minutes (add timeout buffer to 10+ minutes)
- **Storybook Build**: 2-3 minutes (add timeout buffer to 5+ minutes)
- **Browser Install**: 5-10 minutes (add timeout buffer to 15+ minutes)

### Package Management
- Repository uses **Yarn 4.1.0** workspaces with **Lerna 8.1.2**
- Always use `yarn` commands, not `npm`
- Use `--immutable` flag for CI/reproducible installs
- Lockfile modifications may be forbidden in CI environments

### Development Workflow
1. Make changes to TypeScript/SCSS files
2. Watch mode automatically recompiles (`yarn watch`)
3. Test in storybook or dev server
4. Run `yarn lint:fix` before committing
5. Run `yarn build` to ensure production build works
6. Always validate with manual testing scenarios above

## Environment Requirements

- **Node.js**: Version 20.x (current repository uses 20.19.4)
- **Yarn**: Version 4.1.0 (managed by packageManager field)
- **Git**: Required for version management and lerna operations
- **Playwright Browsers**: Required for testing (optional for development)

## Troubleshooting

### "Cannot find module" errors
- Run `yarn build` to ensure all packages are compiled
- Check that dependencies are installed with `yarn install --immutable`

### Test failures
- Ensure browsers are installed: `yarn playwright install chromium --with-deps`
- Check network connectivity for browser downloads
- Tests may not work in restricted environments

### Build errors  
- Run `yarn clean` then `yarn build` for clean rebuild
- Check TypeScript configuration in individual packages
- Ensure Sass files compile with `yarn sass` in mgt-components

### Storybook issues
- Run `yarn storybook:dev` to regenerate custom-elements.json
- Check rollup bundle with `yarn storybook:bundle`
- Clear storybook cache if components don't update