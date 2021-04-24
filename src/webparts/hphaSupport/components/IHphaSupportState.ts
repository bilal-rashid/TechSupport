export interface IHphaSupportState {
  errorConfig: boolean;
  showSuccess: boolean;
  loading: boolean;
  showDataUpload: boolean;
  selectedTitle: string|number;
  selectedScenario: string|number;
  resultRecord: any;
  items: any[];
  uniqueTitles: any[];
  filteredScenario: any[];
  jsonArray: string;
}