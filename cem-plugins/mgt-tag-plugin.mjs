import { resolveModuleOrPackageSpecifier } from '@custom-elements-manifest/analyzer/src/utils/index.js';
export default function mgtTagPlugin() {
  function isCustomRegistration(node) {
    // this would be better if we tested arg[0] for a string literal
    return node?.expression?.getText() === 'registerComponent' && node.arguments.length >= 2;
  }
  // Write a custom plugin
  return {
    // Make sure to always give your plugins a name, this helps when debugging
    name: 'mgt-tag-plugin',
    // Runs for all modules in a project, before continuing to the analyzePhase
    collectPhase({ ts, node, context }) {},
    // Runs for each module
    analyzePhase({ ts: TS, node, moduleDoc, context }) {
      if (isCustomRegistration(node)) {
        //do things...
        if (context.dev) console.log(node);

        var elementTag = node.arguments[0].text;
        var elementClass = node.arguments[1].text;
        const definitionDoc = {
          kind: 'custom-element-definition',
          name: elementTag,
          declaration: {
            name: elementClass,
            ...resolveModuleOrPackageSpecifier(moduleDoc, context, elementClass)
          }
        };
        moduleDoc.exports = [...(moduleDoc.exports || []), definitionDoc];
      }
    },
    // Runs for each module, after analyzing, all information about your module should now be available
    moduleLinkPhase({ moduleDoc, context }) {},
    // Runs after modules have been parsed and after post-processing
    packageLinkPhase({ customElementsManifest, context }) {}
  };
}
