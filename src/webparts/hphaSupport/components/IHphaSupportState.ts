export interface IHphaSupportState {
  errorConfig: boolean;
  showSuccess: boolean;
  loading: boolean;
  showDataUpload: boolean;
  selectedTitle: string|number;
  selectedSecondCategory: string|number;
  selectedThirdCategory: string|number;
  selectedScenario: string|number;
  resultRecord: any;
  items: any[];
  searchResults: any[];
  showSearchResults: boolean;
  uniqueTitles: any[];
  filteredScenario: any[];
  filteredSecondCategory: any[];
  filteredThirdCategory: any[];
  jsonArray: string;
  stringItemsFirst: any[];
  stringItemsSecond: any[];
  stringItemsThird: any[];
  stringItemsFourth: any[];
  stringItemsFifth: any[];
  stringItemsSixth: any[];
  stringItemsSeventh: any[];
}
