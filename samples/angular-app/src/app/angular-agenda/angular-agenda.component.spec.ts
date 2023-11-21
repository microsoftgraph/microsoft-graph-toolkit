import { ComponentFixture, TestBed, waitForAsync } from '@angular/core/testing';
import { AngularAgendaComponent } from './angular-agenda.component';

describe('AngularAgendaComponent', () => {
  let component: AngularAgendaComponent;
  let fixture: ComponentFixture<AngularAgendaComponent>;

  beforeEach(waitForAsync(() => {
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
