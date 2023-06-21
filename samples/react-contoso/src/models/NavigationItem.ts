export interface NavigationItem {
  name: string;
  url: string;
  pattern?: string;
  icon: any;
  key: string;
  requiresLogin: boolean;
  component: any;
  exact: boolean;
}
