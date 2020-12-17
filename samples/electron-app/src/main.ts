import { app, BrowserWindow } from 'electron';
import * as path from 'path';
import { ElectronAuthenticator } from './ElectronAuthenticator';

export default class Main {
  static application: Electron.App;
  public static mainWindow: Electron.BrowserWindow;
  static accessToken: string;
  static authProvider: ElectronAuthenticator;
  //static networkModule: FetchManager;

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
      authority:
        'https://login.microsoftonline.com/29d87d10-93c7-4330-8d09-d6ac77ae5af2',
      mainWindow: Main.mainWindow,
      // scopes: ['User.Read'],
    });
    Main.mainWindow.loadFile(path.join(__dirname, '../index.html'));
    Main.mainWindow.on('closed', Main.onClose);
  }

  private static createMainWindow(): void {
    this.mainWindow = new BrowserWindow({
      width: 800,
      height: 800,
      webPreferences: {
        nodeIntegration: true,
      },
    });
  }

  private static onClose(): void {
    Main.mainWindow = null;
  }
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.

// In this file you can include the rest of your app"s specific main process
// code. You can also put them in separate files and require them here.
