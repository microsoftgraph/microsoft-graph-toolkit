import { Component } from '@angular/core';
import {MgtPersonDetails} from '@microsoft/mgt';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  person : MgtPersonDetails = {
    displayName: "Nikola Metulev"
  };

  loginCompleted = e => {
    console.log('loginCompleted');
  }
}
