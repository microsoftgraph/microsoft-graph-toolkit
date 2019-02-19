import '../../../dist/types/stencil.core';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
export declare class LoginComponent {
    $rootElement: HTMLElement;
    private $menuElement;
    user: MicrosoftGraph.User;
    profileImage: string | ArrayBuffer;
    componentWillLoad(): Promise<void>;
    componentDidLoad(): void;
    componentDidUpdate(): void;
    private init;
    private loadState;
    login(): Promise<void>;
    logout(): Promise<void>;
    toggleMenu(): void;
    clicked(): void;
    render(): JSX.Element;
    renderLoggedOut(): JSX.Element;
    renderLoggedIn(): JSX.Element;
    renderMenu(): JSX.Element;
}
