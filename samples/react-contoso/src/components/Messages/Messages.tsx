import { MgtTemplateProps, Person, PersonCardInteraction, PersonViewType } from '@microsoft/mgt-react';
import './Messages.css';

export function Messages(props: MgtTemplateProps) {
  const email = props.dataContext;
  return (
    <div className="email">
      <a href={email.webLink} target="_blank" rel="noreferrer">
        <div className="header">
          <div>
            <Person
              personQuery={email.sender.emailAddress.address}
              view={PersonViewType.oneline}
              personCardInteraction={PersonCardInteraction.hover}
            />
          </div>
        </div>
        <div className="title">
          <h3>{email.subject}</h3>
          <span className="date">{new Date(email.receivedDateTime).toLocaleDateString()}</span>
        </div>
        {email.bodyPreview ?? <div className="preview">{email.bodyPreview}</div>}
        {!email.bodyPreview ?? <div className="preview empty-body">...</div>}
      </a>
    </div>
  );
}
