import * as React from 'react';
import styles from './HphaSupport.module.scss';
import { IHphaSupportProps } from './IHphaSupportProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { default as pnp } from "sp-pnp-js";
import {IHphaSupportState} from "./IHphaSupportState";
import { DefaultButton, PrimaryButton, IStackTokens } from 'office-ui-fabric-react';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Label } from 'office-ui-fabric-react/lib/Label';
import {  MessageBar,
  MessageBarType } from 'office-ui-fabric-react';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };
const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 400 } };
const stackTokens = { childrenGap: 15 };
const stackTokensDetails = { childrenGap: 3 };

const ListTitle = 'HphaSupport';
// const ListTitle = 'Testtt';
const AdminTitle = 'john.brennan@hpha';
// const AdminTitle = 'bilal.rashid@slickwhiz';
export default class HphaSupport extends React.Component<IHphaSupportProps, IHphaSupportState> {

  public componentDidMount(): void {
    this.setState({errorConfig: false, loading: false});
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
  public onChangeTitle = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    let array = [];
    const filtered = this.state.items.filter(p => p.Title === item.key);
    const unique = new Set(filtered.map(row => row.SpecificScenario));
    unique.forEach(t => {
      array.push({key: t, text: t});
    });
    array.sort((a, b) => {
      if (a.text < b.text) { return -1; }
      if (a.text > b.text) { return 1; }
      return 0;
    });
    this.setState({selectedTitle: item.key, selectedScenario: null,filteredScenario: array,resultRecord:null});
  }
  public onChangeScenario = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    // const filtered = this.state.items.filter(p => p.Title === item.key);
    const filtered = this.state.items.filter(p => p.Title === this.state.selectedTitle && p.SpecificScenario === item.key);
    this.setState({selectedScenario: item.key, resultRecord: (filtered && filtered.length > 0)?filtered[0]:null});
  }
  public getData = () => {
    pnp.sp.web.currentUser.get().then(user => {
      if (user.LoginName.indexOf(AdminTitle) === -1) {
        this.setState({showDataUpload: false});
      } else {
        this.setState({showDataUpload: true});
      }
    });
    pnp.sp.web.lists.getByTitle(ListTitle).items.select(
      "Title","SpecificScenario","FirstTier","SecondTier","TroubleshootingTips","OtherDetails", "Id").
    get().then(items => {
      let array = [];
      const unique = new Set(items.map(item => item.Title));
      unique.forEach(t => {
        array.push({key: t, text: t});
      });
      array.sort((a, b) => {
        if (a.text < b.text) { return -1; }
        if (a.text > b.text) { return 1; }
        return 0;
      });
      this.setState({items: items, uniqueTitles: array.sort(),resultRecord:null,errorConfig:false});
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
  public render(): React.ReactElement<IHphaSupportProps> {
    console.log('STATE',this.state);
    return (
    <Stack tokens={stackTokens}>
      {
        (this.state && !this.state.errorConfig)?
        <div>
          <Stack tokens={stackTokens}>
            <Dropdown
              label={this.props.equipment}
              selectedKey={this.state && this.state.selectedTitle ? this.state.selectedTitle : undefined}
              onChange={this.onChangeTitle}
              placeholder="Select an option"
              options={this.state && this.state.uniqueTitles && this.state.uniqueTitles.length > 0 ? this.state.uniqueTitles : []}
              styles={dropdownStyles}
            />
            <Dropdown
              label={this.props.issues}
              selectedKey={this.state && this.state.selectedScenario ? this.state.selectedScenario : null}
              onChange={this.onChangeScenario}
              placeholder="Select an option"
              options={this.state && this.state.filteredScenario && this.state.filteredScenario.length > 0 ? this.state.filteredScenario : []}
              styles={dropdownStyles}
            />
          </Stack>
          {
            (this.state && this.state.resultRecord)?
              <Stack tokens={stackTokensDetails}>
                <Label disabled>{this.props.tips}</Label>
                <Label style={{fontSize:21}}>{this.state.resultRecord.TroubleshootingTips}</Label>
                <br/>
                <Label disabled>{this.props.first}</Label>
                <Label style={{fontSize:21}}>{this.state.resultRecord.FirstTier}</Label>
                <br/>
                <Label disabled>{this.props.second}</Label>
                <Label style={{fontSize:21}}>{this.state.resultRecord.SecondTier}</Label>
                <br/>
                {(this.state.resultRecord.OtherDetails)&&<Label disabled>Other Details</Label>}
                {(this.state.resultRecord.OtherDetails)&&<Label style={{fontSize:21}}>{this.state.resultRecord.OtherDetails}</Label>}
              </Stack>:null
          }
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
}
