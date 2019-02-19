import '../../../dist/types/stencil.core';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
export declare class AgendaComponent {
    _things: Array<MicrosoftGraph.Event>;
    eventTemplateFunction: (event: any) => string;
    $rootElement: HTMLElement;
    private _provider;
    componentWillLoad(): Promise<void>;
    private init;
    private loadData;
    render(): JSX.Element;
    private renderEvent;
    private renderEventTemplate;
    getStartingTime(event: MicrosoftGraph.Event): string;
    getEventDuration(event: MicrosoftGraph.Event): string;
}
