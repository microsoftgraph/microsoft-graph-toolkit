# @microsoft/mgt-element

Going forward in version 2, Microsoft Graph Toolkit will be broken up into multiple packages. This allows developers to take what they need from the toolkit and avoid any unnecessary parts.

`@microsft/mgt-element` is the first of these packages. It contains the low level interfaces and base classes that all MGT components and providers are built upon.

```ts
<script type="module">
    import { IProvider, Providers } from '@microsoft/mgt-element'

    export class MyProvider extends IProvider {
        // Create your own provider
    }

    const provider: IProvider = new MyProvider();
    Providers.globalProvider = provider;

    const graph: IGraph = provider.graph;
    const me: any = await graph.api('/me').get();
</script>
```