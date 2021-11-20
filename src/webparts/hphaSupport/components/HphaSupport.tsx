import * as React from 'react';
import styles from './HphaSupport.module.scss';
import { IHphaSupportProps } from './IHphaSupportProps';
import { default as pnp } from "sp-pnp-js";
import {IHphaSupportState} from "./IHphaSupportState";
import { PrimaryButton, IIconProps, List} from 'office-ui-fabric-react';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Label } from 'office-ui-fabric-react/lib/Label';
import {  MessageBar,
  MessageBarType } from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { Dropdown, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };
const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 400 } };
const stackTokens = { childrenGap: 15 };
const stackTokensDetails = { childrenGap: 3 };
const refreshIcon: IIconProps = { iconName: 'Refresh' };
// const ListTitle = 'MyList';
const ListTitle = 'TechnicalSupport';
const AdminTitle = 'john.brennan@hpha';
// const AdminTitle = 'bilal.rashid@slickwhiz';
// old cdn https://hpeits.sharepoint.com/sites/HPHAAMGHSupport/SiteAssets/technicalsupport/
var customStyles = {
  tableHeader:{background: '#5B9BD5',
    display: 'table-cell',
    padding: 5,
    width:200,
    color:'#fff',
    fontWeight:700,
    borderRight: '1px solid #5B9BD5',
    borderLeft: '1px solid #5B9BD5'},
  tableHeader2:{background: '#5B9BD5',
    display: 'table-cell',
    padding: 5,
    width:399,
    color:'#fff',
    fontWeight:700,
    borderRight: '1px solid #5B9BD5',
    borderLeft: '1px solid #5B9BD5'},
  colorTableCellKey: {background: '#DEEAF6',
    display: 'table-cell',
    padding: 5,
    width:200,
    fontWeight:600,
    borderRight: '1px solid #5B9BD5',
    borderLeft: '1px solid #5B9BD5'},
  colorTableCellValue: {background: '#DEEAF6',
    display: 'table-cell',
    padding: 5,
    width:400,
    fontWeight:600,
    borderRight: '1px solid #5B9BD5'},
  whiteTableCellKey: {background: '#fff',
    display: 'table-cell',
    padding: 5,
    width:200,
    fontWeight:600,
    borderRight: '1px solid #5B9BD5',
    borderLeft: '1px solid #5B9BD5'},
  whiteTableCellValue: {background: '#fff',
    display: 'table-cell',
    padding: 5,
    width:400,
    fontWeight:600,
    borderRight: '1px solid #5B9BD5'},
  line:{height:1,width:623,
    background: '#5B9BD5'}
};
export default class HphaSupport extends React.Component<IHphaSupportProps, IHphaSupportState> {

  public onRenderCell = (item: any, index: number | undefined): JSX.Element => {
    return (
      <div style={{display:'table', marginTop:20}}>
        <div>
          <div style={customStyles.tableHeader}>
            {index+1}
          </div>
          <div style={customStyles.tableHeader2}>
          </div>
        </div>
        <div>
          <div style={customStyles.colorTableCellKey}>
            {this.props.firstCategory}
          </div>
          <div style={customStyles.colorTableCellValue}>
            {item.Title}
          </div>
        </div>
        <div>
          <div style={customStyles.colorTableCellKey}>
            {this.props.secondCategory}
          </div>
          <div style={customStyles.colorTableCellValue}>
            {item.SecondaryCategory}
          </div>
        </div>
        <div>
          <div style={customStyles.colorTableCellKey}>
            {this.props.thirdCategory}
          </div>
          <div style={customStyles.colorTableCellValue}>
            {item.ThirdCategory}
          </div>
        </div>
        <div>
          <div style={customStyles.colorTableCellKey}>
            {this.props.issues}
          </div>
          <div style={customStyles.colorTableCellValue}>
            {item.SpecificIssue}
          </div>
        </div>
        <div style={customStyles.line}></div>
        <div>
          <div style={customStyles.whiteTableCellKey}>
            {this.props.tips}
          </div>
          <div style={customStyles.whiteTableCellValue}>
            {item.TroubleshootingTips}
          </div>
        </div>
        <div>
          <div style={customStyles.whiteTableCellKey}>
            {this.props.firstSupport}
          </div>
          <div style={customStyles.whiteTableCellValue}>
            {item.FirstTierSupport}
          </div>
        </div>
        {/*<div>*/}
        {/*  <div style={customStyles.whiteTableCellKey}>*/}
        {/*    {this.props.secondSupport}*/}
        {/*  </div>*/}
        {/*  <div style={customStyles.whiteTableCellValue}>*/}
        {/*    {item.SecondTierSupport}*/}
        {/*  </div>*/}
        {/*</div>*/}
        <div style={customStyles.line}></div>
        {/*<div>*/}
        {/*  <div style={customStyles.colorTableCellKey}>*/}
        {/*    {this.props.link}*/}
        {/*  </div>*/}
        {/*  <div style={customStyles.colorTableCellValue}>*/}
        {/*    {(item.LinkToSupportMaterial && this.isValidHttpUrl(item.LinkToSupportMaterial)) ?*/}
        {/*      <a href={item.LinkToSupportMaterial} target={'_blank'}>{item.LinkToSupportMaterial}</a>:*/}
        {/*    null}*/}
        {/*  </div>*/}
        {/*</div>*/}
        {/*<div style={customStyles.line}></div>*/}
      </div>
    );
  }
  public componentDidMount(): void {
    this.setState({errorConfig: false, loading: false, showSearchResults:false,showSuccess:false});
    this.getData();
  }
  public populteData = () => {
    this.setState({loading:true});
    const dta = JSON.parse(this.state.jsonArray);
    const len = dta.length;
    let counter = 0;
    if (dta && dta.length > 0) {
      dta.forEach(item => {
        pnp.sp.web.lists.getByTitle(ListTitle).items.add(item).then((iar) => {
          counter++;
          if(counter > len-1) {
            this.showToast();
          }
        }).catch(error=>{
        });
      });
    }
  }
  public onChangeFirstCategory = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item.key && item.key !== this.state.selectedTitle) {
      let array = [];
      const filtered = this.state.items.filter(p => p.Title === item.key && p.SecondaryCategory);
      const unique = new Set(filtered.map(row => row.SecondaryCategory));
      unique.forEach(t => {
        array.push({key: t, text: t});
      });
      array.sort((a, b) => {
        if (a.text < b.text) { return -1; }
        if (a.text > b.text) { return 1; }
        return 0;
      });
      this.setState({showSearchResults:false,selectedTitle: item.key, selectedSecondCategory: null, filteredSecondCategory:[],
      selectedThirdCategory:null, filteredThirdCategory:null,selectedScenario:null, filteredScenario:null,resultRecord:null});
      setTimeout(() => {  this.setState({filteredSecondCategory:array});},  100);
    }
  }
  public componentDidUpdate(prevProps) {
    if (prevProps.colorLightBackground !== this.props.colorLightBackground ||
      prevProps.colorBackground !== this.props.colorBackground ||
      prevProps.colorHeader !== this.props.colorHeader) {
      let searchItems = this.state.searchResults.filter(p => p.Id !== -1);
      this.setState({searchResults: searchItems});
    }
  }
  public onChangeSecondCategory = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item.key && item.key !== this.state.selectedSecondCategory) {
      let array = [];
      const filtered = this.state.items.filter(p => p.Title === this.state.selectedTitle && p.SecondaryCategory === item.key && p.ThirdCategory);
      console.log(filtered);
      const unique = new Set(filtered.map(row => row.ThirdCategory));
      console.log(unique);
      unique.forEach(t => {
        array.push({key: t, text: t});
      });
      array.sort((a, b) => {
        if (a.text < b.text) { return -1; }
        if (a.text > b.text) { return 1; }
        return 0;
      });
      this.setState({selectedSecondCategory: item.key,
        selectedThirdCategory:null, filteredThirdCategory:[],selectedScenario:null,filteredScenario:null,resultRecord:null});
      setTimeout(() => {  this.setState({filteredThirdCategory:array});},  100);
    }
  }
  public onChangeThirdCategory = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item.key  && item.key !== this.state.selectedThirdCategory) {
      let array = [];
      const filtered = this.state.items.filter(p => p.Title === this.state.selectedTitle && p.SecondaryCategory === this.state.selectedSecondCategory &&
        p.ThirdCategory === item.key && p.SpecificIssue);
      console.log(filtered);
      const unique = new Set(filtered.map(row => row.SpecificIssue));
      console.log(unique);
      unique.forEach(t => {
        array.push({key: t, text: t});
      });
      array.sort((a, b) => {
        if (a.text < b.text) { return -1; }
        if (a.text > b.text) { return 1; }
        return 0;
      });
      this.setState({selectedThirdCategory:item.key,
        selectedScenario:null,filteredScenario:[],resultRecord:null});
      setTimeout(() => {  this.setState({filteredScenario:array});},  100);
    }
  }
  public onChangeScenario = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item.key) {
      const filtered = this.state.items.filter(p => p.Title === this.state.selectedTitle && p.SecondaryCategory === this.state.selectedSecondCategory &&
        p.ThirdCategory === this.state.selectedThirdCategory && p.SpecificIssue === item.key);
      this.setState({selectedScenario: item.key, resultRecord: (filtered && filtered.length > 0)?filtered[0]:null});
    }
  }
  public getData = () => {
    pnp.sp.web.currentUser.get().then(user => {
      if (user.LoginName.indexOf(AdminTitle) === -1) {
        this.setState({showDataUpload: false});
      } else {
        this.setState({showDataUpload: true});
      }
    });
    pnp.sp.web.lists.getByTitle(ListTitle).items.top(4000).select(
      "Title","SecondaryCategory","ThirdCategory","SpecificIssue","TroubleshootingTips","FirstTierSupport","SecondTierSupport","Tags", "Id").
    get().then(items => {
      let array = [];
      const unique = new Set(items.map(item => item.Title));
      unique.forEach(t => {
        if (t) {
          array.push({key: t, text: t});
        }
      });
      array.sort((a, b) => {
        if (a.text < b.text) { return -1; }
        if (a.text > b.text) { return 1; }
        return 0;
      });
      let stringArrayFirst = [];
      let stringArraySecond = [];
      let stringArrayThird = [];
      let stringArrayFourth = [];
      let stringArrayFifth = [];
      let stringArraySixth = [];
      let stringArraySeventh = [];
      if (items && items.length > 0) {
        items.forEach(item => {
          stringArrayFirst.push({id: item.Id ,text: '' + item.Title});
          stringArraySecond.push({id: item.Id ,text: '' + item.SecondaryCategory});
          stringArrayThird.push({id: item.Id ,text: '' + item.ThirdCategory});
          stringArrayFourth.push({id: item.Id ,text: '' + item.SpecificIssue});
          stringArrayFifth.push({id: item.Id ,text: '' + item.TroubleshootingTips});
          stringArraySixth.push({id: item.Id ,text: '' + item.Tags});
          stringArraySeventh.push({id: item.Id ,text: '' + item.FirstTierSupport + '   '
              + item.SecondTierSupport});
        });
      }
      this.setState({
        stringItemsFirst: stringArrayFirst,
        stringItemsSecond: stringArraySecond,
        stringItemsThird: stringArrayThird,
        stringItemsFourth: stringArrayFourth,
        stringItemsFifth: stringArrayFifth,
        stringItemsSixth: stringArraySixth,
        stringItemsSeventh: stringArraySeventh,
        items: items, uniqueTitles: array,resultRecord:null,errorConfig:false});
    }).catch(error => {
      this.setState({errorConfig: true});
    });
  }
  public showToast = () => {
    this.setState({loading:false, jsonArray:''});
    this.setState({showSuccess: true});
    setTimeout(() => {
      this.setState({showSuccess: false});
    }, 2000);
  }
  public onChangeFirstTextFieldValue =
    (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    this.setState({jsonArray: newValue || ''});
  }
  public onSearch = (newValue) => {
    if (newValue && newValue.length > 1) {
      const results1 = this.state.stringItemsFirst.filter(p => p.text.toLowerCase().indexOf(newValue.toLowerCase()) !== -1);
      const results2 = this.state.stringItemsSecond.filter(p => p.text.toLowerCase().indexOf(newValue.toLowerCase()) !== -1);
      const results3 = this.state.stringItemsThird.filter(p => p.text.toLowerCase().indexOf(newValue.toLowerCase()) !== -1);
      const results4 = this.state.stringItemsFourth.filter(p => p.text.toLowerCase().indexOf(newValue.toLowerCase()) !== -1);
      const results5 = this.state.stringItemsFifth.filter(p => p.text.toLowerCase().indexOf(newValue.toLowerCase()) !== -1);
      const results6 = this.state.stringItemsSixth.filter(p => p.text.toLowerCase().indexOf(newValue.toLowerCase()) !== -1);
      const results7 = this.state.stringItemsSeventh.filter(p => p.text.toLowerCase().indexOf(newValue.toLowerCase()) !== -1);
      let ids = [];
      results1.forEach(result => (ids.indexOf(result.id) === -1) ? ids.push(result.id):false);
      results2.forEach(result => (ids.indexOf(result.id) === -1) ? ids.push(result.id):false);
      results3.forEach(result => (ids.indexOf(result.id) === -1) ? ids.push(result.id):false);
      results4.forEach(result => (ids.indexOf(result.id) === -1) ? ids.push(result.id):false);
      results5.forEach(result => (ids.indexOf(result.id) === -1) ? ids.push(result.id):false);
      results6.forEach(result => (ids.indexOf(result.id) === -1) ? ids.push(result.id):false);
      results7.forEach(result => (ids.indexOf(result.id) === -1) ? ids.push(result.id):false);
      console.log('IDs', ids);
      let searchItems = [];
      ids.forEach(id => {
        const resultItem = this.state.items.filter(p => p.Id === id);
        if (resultItem && resultItem.length > 0) {
          searchItems.push(resultItem[0]);
        }
      });
      this.setState({showSearchResults:true,searchResults: searchItems,selectedTitle: null, selectedSecondCategory: null, filteredSecondCategory:[],
        selectedThirdCategory:null, filteredThirdCategory:null,selectedScenario:null, filteredScenario:null,resultRecord:null});
    } else {
      this.setState({showSearchResults:false});
    }
  }
  public refreshStyles = () => {
    customStyles = {
      tableHeader:{background: this.props.colorHeader,
        display: 'table-cell',
        padding: 5,
        width:200,
        color:'#fff',
        fontWeight:700,
        borderRight: '1px solid '+this.props.colorHeader,
        borderLeft: '1px solid '+this.props.colorHeader},
      tableHeader2:{background: this.props.colorHeader,
        display: 'table-cell',
        padding: 5,
        width:399,
        color:'#fff',
        fontWeight:700,
        borderRight: '1px solid '+this.props.colorHeader,
        borderLeft: '1px solid '+this.props.colorHeader},
      colorTableCellKey: {background: this.props.colorBackground,
        display: 'table-cell',
        padding: 5,
        width:200,
        fontWeight:600,
        borderRight: '1px solid '+this.props.colorHeader,
        borderLeft: '1px solid '+this.props.colorHeader},
      colorTableCellValue: {background: this.props.colorBackground,
        display: 'table-cell',
        padding: 5,
        width:400,
        fontWeight:600,
        borderRight: '1px solid '+this.props.colorHeader},
      whiteTableCellKey: {background: this.props.colorLightBackground,
        display: 'table-cell',
        padding: 5,
        width:200,
        fontWeight:600,
        borderRight: '1px solid '+this.props.colorHeader,
        borderLeft: '1px solid '+this.props.colorHeader},
      whiteTableCellValue: {background: this.props.colorLightBackground,
        display: 'table-cell',
        padding: 5,
        width:400,
        fontWeight:600,
        borderRight: '1px solid '+this.props.colorHeader},
      line:{height:1,width:623,
        background: this.props.colorHeader}
    };
  }
  public render(): React.ReactElement<IHphaSupportProps> {
    console.log('STATE',this.state);
    this.refreshStyles();
    return (
    <Stack tokens={stackTokens}>
      {
        (this.state && !this.state.errorConfig)?
        <div>
          <Stack horizontal={true} tokens={{ childrenGap: 40 }}>
            {/*<Stack horizontal={false} tokens={stackTokens}>*/}
            {/*  <Label style={{fontWeight:'bold', fontSize:14}}>Use Fields Below to narrow down your issue</Label>*/}
            {/*  <Dropdown*/}
            {/*    notifyOnReselect = {true}*/}
            {/*    label={this.props.firstCategory}*/}
            {/*    selectedKey={this.state && this.state.selectedTitle ? this.state.selectedTitle : undefined}*/}
            {/*    onChange={this.onChangeFirstCategory}*/}
            {/*    placeholder="Select an option"*/}
            {/*    options={this.state && this.state.uniqueTitles && this.state.uniqueTitles.length > 0 ? this.state.uniqueTitles : []}*/}
            {/*    styles={dropdownStyles}*/}
            {/*  />*/}
            {/*  {(this.state && this.state.filteredSecondCategory && this.state.filteredSecondCategory.length > 0)&&<Dropdown*/}
            {/*    notifyOnReselect = {true}*/}
            {/*    label={this.props.secondCategory}*/}
            {/*    selectedKey={this.state && this.state.selectedSecondCategory ? this.state.selectedSecondCategory : undefined}*/}
            {/*    onChange={this.onChangeSecondCategory}*/}
            {/*    placeholder="Select an option"*/}
            {/*    options={this.state && this.state.filteredSecondCategory && this.state.filteredSecondCategory.length > 0 ? this.state.filteredSecondCategory : []}*/}
            {/*    styles={dropdownStyles}*/}
            {/*  />}*/}
            {/*  {(this.state && this.state.filteredThirdCategory && this.state.filteredThirdCategory.length > 0 )&&<Dropdown*/}
            {/*    notifyOnReselect = {true}*/}
            {/*    label={this.props.thirdCategory}*/}
            {/*    selectedKey={this.state && this.state.selectedThirdCategory ? this.state.selectedThirdCategory : undefined}*/}
            {/*    onChange={this.onChangeThirdCategory}*/}
            {/*    placeholder="Select an option"*/}
            {/*    options={this.state && this.state.filteredThirdCategory && this.state.filteredThirdCategory.length > 0 ? this.state.filteredThirdCategory : []}*/}
            {/*    styles={dropdownStyles}*/}
            {/*  />}*/}
            {/*  {(this.state && this.state.filteredScenario && this.state.filteredScenario.length > 0)&&<Dropdown*/}
            {/*    notifyOnReselect = {true}*/}
            {/*    label={this.props.issues}*/}
            {/*    selectedKey={this.state && this.state.selectedScenario ? this.state.selectedScenario : null}*/}
            {/*    onChange={this.onChangeScenario}*/}
            {/*    placeholder="Select an option"*/}
            {/*    options={this.state && this.state.filteredScenario && this.state.filteredScenario.length > 0 ? this.state.filteredScenario : []}*/}
            {/*    styles={dropdownStyles}*/}
            {/*  />}*/}
            {/*</Stack>*/}
            {/*<Label style={{fontWeight:'bold', fontSize:17, marginTop:70}}>OR</Label>*/}
            <Stack horizontal={false} tokens={stackTokens}>
              <Label style={{fontWeight:'bold', fontSize:14}}>Do a keyword search to find a support solution:</Label>
              <div style={{marginTop:20}}>
                <SearchBox style={{width:270}} placeholder="Search" onSearch={this.onSearch} />
              </div>
            </Stack>
          </Stack>
          {
            (this.state && this.state.resultRecord)?
              <Stack tokens={stackTokensDetails}>
                <Label disabled>{this.props.tips}</Label>
                <Label style={{fontSize:21}}>{this.state.resultRecord.TroubleshootingTips}</Label>
                <br/>
                <Label disabled>{this.props.firstSupport}</Label>
                <Label style={{fontSize:21}}>{this.state.resultRecord.FirstTierSupport}</Label>
                <br/>
                {/*<Label disabled>{this.props.secondSupport}</Label>*/}
                {/*<Label style={{fontSize:21}}>{this.state.resultRecord.SecondTierSupport}</Label>*/}
                {/*<br/>*/}
                {/*{(this.state.resultRecord.LinkToSupportMaterial && this.isValidHttpUrl(this.state.resultRecord.LinkToSupportMaterial)) && <Label disabled>{this.props.link}</Label>}*/}
                {/*{(this.state.resultRecord.LinkToSupportMaterial && this.isValidHttpUrl(this.state.resultRecord.LinkToSupportMaterial)) && <Label><a target={'_blank'} href={this.state.resultRecord.LinkToSupportMaterial}>{this.state.resultRecord.LinkToSupportMaterial}</a ></Label>}*/}
                {/*{(this.state.resultRecord.OtherDetails)&&<Label disabled>Other Details</Label>}*/}
                {/*{(this.state.resultRecord.OtherDetails)&&<Label style={{fontSize:21}}>{this.state.resultRecord.OtherDetails}</Label>}*/}
              </Stack>:null
          }
          {(this.state && this.state.showSearchResults) &&
          <Stack tokens={stackTokens}>
            {(this.state.searchResults.length > 0) ?
              <List
                onShouldVirtualize = {()=>false}
                items={this.state.searchResults}
                onRenderCell={this.onRenderCell} /> :
              <div style={{marginTop:20, width:300}}>
                <MessageBar
                  messageBarType={MessageBarType.error}>
                  Your search did not match any results
                </MessageBar>
              </div>}
          </Stack>}
          {
            (this.state && this.state.showDataUpload) &&
            <Stack tokens={stackTokens}>
              <br/>
              <br/>
              <br/>
              <TextField
                multiline
                label="Input Data"
                value={(this.state && this.state.jsonArray)?this.state.jsonArray:''}
                onChange={this.onChangeFirstTextFieldValue}
                styles={textFieldStyles}
              />
              <PrimaryButton disabled={this.state.loading} style={{width: 400}} text="Submit" onClick={this.populteData}  />
              {
                (this.state && this.state.showSuccess) &&
                <MessageBar
                  messageBarType={MessageBarType.success}>
                  Data uploaded successfully
                </MessageBar>
              }
            </Stack>
          }
        </div>:
        <MessageBar
        messageBarType={MessageBarType.error}>
        Error fetching data from list, please create a valid list with Title ({ListTitle})
        </MessageBar>
      }
    </Stack>
    );
  }
  public isValidHttpUrl = (string) => {
    let url;

    try {
      url = new URL(string);
    } catch (_) {
      return false;
    }

    return url.protocol === "http:" || url.protocol === "https:";
  }
}
