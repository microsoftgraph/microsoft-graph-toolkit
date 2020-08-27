import { async, ComponentFixture, TestBed } from '@angular/core/testing';
import { AngularAgendaComponent } from './angular-agenda.component';

describe('AngularAgendaComponent', () => {
  let component: AngularAgendaComponent;
  let fixture: ComponentFixture<AngularAgendaComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ AngularAgendaComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(AngularAgendaComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
