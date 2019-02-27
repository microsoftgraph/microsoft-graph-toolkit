import '../../../dist/types/stencil.core';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
export declare class PersonaComponent {
    imageSize: number;
    id: string;
    function(newValue: string): void;
    persona: GraphPersonaUser;
    user: MicrosoftGraph.User;
    profileImage: string;
    componentWillLoad(): Promise<void>;
    private loadImage;
    render(): JSX.Element;
    renderImage(): JSX.Element;
    renderEmptyImage(): JSX.Element;
    getInitials(user: MicrosoftGraph.User): string;
}
export declare interface GraphPersonaUser {
    displayName: string;
    image?: string;
}
