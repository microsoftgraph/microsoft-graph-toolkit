import { Component } from '@angular/core';
import {MgtPersonDetails} from 'microsoft-graph-toolkit';
import 'microsoft-graph-toolkit/dist/es6/components/mgt-person/mgt-person';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  person : MgtPersonDetails = {
    displayName: "Nikola Metulev"
  };
}
