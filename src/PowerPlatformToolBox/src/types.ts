export interface TemplateRecord {
  documenttemplateid: string; name: string; associatedentitytypecode: string;
  documenttype: number; languagecode?: number; content?: string; clientdata?: string;
  description?: string; createdon?: string; modifiedon?: string;
}
export interface DeploymentResult { target: string; template: string; action: string; success: boolean; bindings: number | null; message: string; }
