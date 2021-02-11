import { app, BrowserWindow } from 'electron';
import * as path from 'path';
import {
  ElectronAuthenticator,
  SimpleCachePlugin,
  MsalElectronConfig
} from '@microsoft/mgt-electron-provider/dist/Authenticator';
export default class Main {
  static application: Electron.App;
  public static mainWindow: Electron.BrowserWindow;

  static main(): void {
    Main.application = app;
    Main.application.on('window-all-closed', Main.onWindowAllClosed);
    Main.application.on('ready', Main.onReady);
  }
  static onWindowAllClosed(arg0: string, onWindowAllClosed: any) {
    Main.application.quit();
  }

  static onReady() {
    Main.createMainWindow();
    let config: MsalElectronConfig = {
      clientId: 'e5774a5b-d84d-4f30-892e-c4f73a503943',
      // authority: '<authority URL optional, uses common authority by default>',
      mainWindow: Main.mainWindow,
      scopes: [
        'user.read',
        'people.read',
        'user.readbasic.all',
        'contacts.read',
        'calendars.read',
        'presence.read.all',
        'tasks.readwrite',
        'presence.read',
        'user.read.all',
        'group.read.all',
        'tasks.read'
      ],
      //SimpleCachePlugin is used here as an example. It stores unencrypted tokens directly on your machine, so please only use it for testing purposes
      cachePlugin: SimpleCachePlugin
    };
    //Initialize the authenticator
    ElectronAuthenticator.initialize(config);
    Main.mainWindow.loadFile(path.join(__dirname, '../index.html'));
    Main.mainWindow.on('closed', Main.onClose);
  }

  private static createMainWindow(): void {
    this.mainWindow = new BrowserWindow({
      width: 800,
      height: 800,
      webPreferences: {
        nodeIntegration: true
      }
    });
  }

  private static onClose(): void {
    Main.mainWindow = null;
  }
}
