# @microsoft/mgt-element

Going forward in version 2, Microsoft Graph Toolkit will be broken up into multiple packages. This allows developers to take what they need from the toolkit and avoid any unnecessary parts.

`@microsft/mgt-element` is the first of these packages. It holds the most low level interfaces and base classes that all MGT components and providers are built upon.

The most notable change for consuming apps is the relocation of the `Providers` class. Access to the global provider instance is managed through `Providers`, so any module references will need to import from `@microsoft/mgt-element` instead.

```ts
<script type="module">
    import { MsalProvider } from '@microsoft/mgt';
    import { IGraph, IProvider, Providers } from '@microsoft/mgt-element'

    const provider: IProvider = new MsalProvider({ clientId: '[CLIENT-ID]' });
    Providers.globalProvider = provider;

    const graph: IGraph = provider.graph;
    const me: any = await graph.api('/me').get();
</script>
```