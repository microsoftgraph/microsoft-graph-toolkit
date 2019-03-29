import {
  LitElement,
  customElement,
  html,
  unsafeCSS,
  property
} from "lit-element";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { Providers } from "../../../library/Providers";
import styles from "./mgt-person.scss";

export interface MgtPersonDetails {
  displayName?: string;
  email?: string;
  image?: string;
  givenName?: string;
  surname?: string;
}

@customElement("mgt-person")
export class MgtPerson extends LitElement {
  @property({ attribute: "image-size" }) imageSize: number = 24;
  @property({ attribute: "show-name", type: Boolean })
  showName: boolean = false;
  @property({ attribute: "show-email", type: Boolean })
  showEmail: boolean = false;
  @property({ attribute: "person-query" }) personQuery: string;
  @property({ attribute: "person-details" }) personDetails: MgtPersonDetails;

  public static get styles() {
    return unsafeCSS(styles);
  }

  public constructor() {
    super();
    Providers.onProviderChange(() => this.onProviderChanged());
    this.loadImage();
  }

  private onProviderChanged() {
    let provider = Providers.getAvailable();
    if (provider.isLoggedIn) this.loadImage();

    provider.onLoginChanged(() => this.loadImage());
  }

  public attributeChangedCallback(name: string, oldVal: any, newVal: any) {
    super.attributeChangedCallback(name, oldVal, newVal);
    if (name === "person-query" && oldVal !== newVal) {
      this.personDetails = null;
      this.loadImage();
    }
  }

  private async loadImage() {
    if (!this.personDetails && this.personQuery) {
      let provider = Providers.getAvailable();

      if (provider && provider.isLoggedIn) {
        if (this.personQuery == "me") {
          let person: MgtPersonDetails = {};

          await Promise.all([
            provider.graph.me().then(user => {
              if (user) {
                person.displayName = user.displayName;
                person.email = user.mail;
              }
            }),
            provider.graph.myPhoto().then(photo => {
              if (photo) {
                person.image = photo;
              }
            })
          ]);

          this.personDetails = person;
        } else {
          provider.graph.findPerson(this.personQuery).then(people => {
            if (people && people.length > 0) {
              let person = people[0] as MicrosoftGraph.Person;
              this.personDetails = person;

              if (
                person.scoredEmailAddresses &&
                person.scoredEmailAddresses.length
              ) {
                this.personDetails.email =
                  person.scoredEmailAddresses[0].address;
              } else if (
                (<any>person).emailAddresses &&
                (<any>person).emailAddresses.length
              ) {
                // beta endpoind uses emailAddresses instead of scoredEmailAddresses
                this.personDetails.email = (<any>(
                  person
                )).emailAddresses[0].address;
              }

              if (person.userPrincipalName) {
                let userPrincipalName = person.userPrincipalName;
                provider.graph.getUserPhoto(userPrincipalName).then(photo => {
                  this.personDetails.image = photo;
                  this.requestUpdate();
                });
              }
            }
          });
        }
      } else {
        this.personDetails = null;
      }
    }
  }

  public render() {
    return html`
      <div class="root">
        ${this.renderImage()} ${this.renderNameAndEmail()}
      </div>
    `;
  }

  private renderImage() {
    if (this.personDetails) {
      if (this.personDetails.image) {
        return html`
          <img
            class="user-avatar ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
            src=${this.personDetails.image as string}
          />
        `;
      } else {
        return html`
          <div
            class="user-avatar initials ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
          >
            <style>
              .initials-text {
                font-size: ${this.imageSize * 0.45}px;
              }
            </style>
            <span class="initials-text">
              ${this.getInitials()}
            </span>
          </div>
        `;
      }
    }
  }
  private renderEmptyImage() {
    return html`
      <i
        class="ms-Icon ms-Icon--Contact avatar-icon ${this.getImageRowSpanClass()} ${this.getImageSizeClass()}"
      ></i>
    `;
  }
  private renderNameAndEmail() {
    if (!this.personDetails || (!this.showName && !this.showEmail)) return;

    let pd = this.personDetails;

    const nameView = this.showName
      ? html`
          <div class="user-name">${pd.displayName}</div>
        `
      : null;

    const emailView = this.showEmail
      ? html`
          <div class="user-email">${pd.email}</div>
        `
      : null;

    return html`
      ${nameView} ${emailView}
    `;
  }

  private getInitials() {
    if (!this.personDetails) return "";

    let pd = this.personDetails || {};
    let ret = "";

    if (pd.givenName) ret += pd.givenName[0].toUpperCase();
    if (pd.surname) ret += pd.surname[0].toUpperCase();

    if (!ret && pd.displayName) {
      let name = pd.displayName.split(" ");
      for (let subName of name) ret += subName[0].toUpperCase();
    }

    return ret;
  }
  private getImageRowSpanClass() {
    return this.showName && this.showEmail ? "row-span-2" : "";
  }
  private getImageSizeClass() {
    return this.showName && this.showEmail ? "small" : "";
  }
}
