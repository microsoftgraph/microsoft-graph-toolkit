export interface Incident {
  id: number;
  title: string;
  description: string;
  status: string;
  priority: string;
  created: Date;
  closed: Date;
  closedBy: string;
  createdBy: string;
  assignedTo: string;
  siteId: string;
  planId: string;
  conversationId: string;
}
