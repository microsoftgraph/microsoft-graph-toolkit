import { app, BrowserWindow } from 'electron';
import * as path from 'path';
import { ElectronAuthenticator } from '@microsoft/mgt-electron-provider/dist/Authenticator';
export default class Main {
  static application: Electron.App;
  public static mainWindow: Electron.BrowserWindow;
  static authProvider: ElectronAuthenticator;

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
    Main.authProvider = new ElectronAuthenticator({
      clientId: 'e5774a5b-d84d-4f30-892e-c4f73a503943',
      authority: 'https://login.microsoftonline.com/29d87d10-93c7-4330-8d09-d6ac77ae5af2',
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
      ]
    });
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
