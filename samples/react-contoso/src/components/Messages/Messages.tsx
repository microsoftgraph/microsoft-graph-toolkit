import { MgtTemplateProps, Person, PersonCardInteraction, PersonViewType } from '@microsoft/mgt-react';
import './Messages.css';

export function Messages(props: MgtTemplateProps) {
  const email = props.dataContext;
  return (
    <div className="email">
      <a href={email.webLink} target="_blank" rel="noreferrer">
        <div className="email-header">
          <div>
            <Person
              personQuery={email.sender.emailAddress.address}
              view={PersonViewType.oneline}
              personCardInteraction={PersonCardInteraction.hover}
            />
          </div>
        </div>
        <div className="email-title">
          <h3>{email.subject}</h3>
          <span className="email-date">{new Date(email.receivedDateTime).toLocaleDateString()}</span>
        </div>
        {email.bodyPreview ?? <div className="email-preview">{email.bodyPreview}</div>}
        {!email.bodyPreview ?? <div className="email-preview email-empty-body">...</div>}
      </a>
    </div>
  );
}
