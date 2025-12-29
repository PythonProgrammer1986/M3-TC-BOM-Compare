
export interface TCBOMRow {
  Lvl: string;
  'Part/PA Id Indented': string;
  'Parent Id': string;
  'Qty'?: string;
  'UOM'?: string;
  [key: string]: string | undefined;
}

export interface M3BOMRow {
  'Rel Lvl': string;
  'Part no': string;
  'Prod struct': string;
  'Qty'?: string;
  'UOM'?: string;
  [key: string]: string | undefined;
}

export interface IIMRow {
  'Item number': string;
  'Planning Policy Name': string;
  'Status'?: string;
  [key: string]: string | undefined;
}

export interface BOMError {
  id: string;
  task: string;
  errorType: 'Missing Level 1' | 'Phantom Mismatch' | 'Policy Violation' | 'Quantity Mismatch' | 'UOM Conflict';
  parentId: string;
  partId: string;
  description: string;
  priority: 'High' | 'Medium' | 'Low';
  actionableFix: string;
}

export interface ValidationSummary {
  totalErrors: number;
  task1Errors: number;
  task2Errors: number;
  task3Errors: number;
  quantityErrors: number;
  status: 'Idle' | 'Validating' | 'Success' | 'Error';
  healthScore: number;
}

export interface TCConnection {
  url: string;
  user: string;
  pass: string;
  itemId: string;
}
