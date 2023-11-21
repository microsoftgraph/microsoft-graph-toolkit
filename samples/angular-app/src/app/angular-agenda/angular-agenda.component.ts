import { AfterViewInit, Component, ElementRef, ViewChild } from '@angular/core';
import { MgtAgenda } from '@microsoft/mgt-components';

@Component({
  selector: 'app-angular-agenda',
  templateUrl: './angular-agenda.component.html',
  styleUrls: ['./angular-agenda.component.scss']
})
export class AngularAgendaComponent implements AfterViewInit {
  @ViewChild('myagenda', { static: true })
  agendaElement: ElementRef<MgtAgenda>;

  constructor() {}

  ngAfterViewInit() {
    this.agendaElement.nativeElement.templateContext = {
      openWebLink: (e: any, context: { event: { webLink: string | undefined } }, root: any) => {
        window.open(context.event.webLink, '_blank');
      },
      getDate: (dateString: string) => {
        const dateObject = new Date(dateString);
        return dateObject.setHours(0, 0, 0, 0);
      },
      getTime: (dateString: string) => {
        const dateObject = new Date(dateString);
        return (
          dateObject.getHours().toString().padEnd(2, '0') + ':' + dateObject.getMinutes().toString().padEnd(2, '0')
        );
      }
    };
  }
}
