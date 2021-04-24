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

// const ListTitle = 'HphaSupport';
const ListTitle = 'Testtt';
const AdminTitle = 'john.brennan@hpha';
// const AdminTitle = 'bilal.rashid@slickwhiz';
const Data = '[{"Title":"GE Monitors","SpecificScenario":"Physical Damage","FirstTier":"Biomed","SecondTier":"GE Canada","ThirdTier":"","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"Alarms not working","FirstTier":"Biomed","SecondTier":"GE Canada","ThirdTier":"","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"Vitals not crossing to Meditech","FirstTier":"IT","SecondTier":"Biomed","ThirdTier":"GE Canada","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"CIC not seeing all montiors","FirstTier":"Biomed","SecondTier":"IT","ThirdTier":"GE Canada","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"Wifi Monitors not transmitting/visible","FirstTier":"IT","SecondTier":"Biomed","ThirdTier":"GE Canada","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"GE strip printer not working","FirstTier":"Biomed","SecondTier":"IT","ThirdTier":"GE Canada","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"GE laser printer not working","FirstTier":"Biomed","SecondTier":"IT","ThirdTier":"GE Canada","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"Strip paper is out","FirstTier":"Materials management for re-order","SecondTier":"Biomed","ThirdTier":"Biomed","OtherDetails":"Not part of the Ricoh toner program"},{"Title":"GE Monitors","SpecificScenario":"Laser printer out of toner","FirstTier":"Materials management for re-order","SecondTier":"Biomed","ThirdTier":"Biomed","OtherDetails":""},{"Title":"GE Monitors","SpecificScenario":"New Monitor","FirstTier":"Biomed","SecondTier":"IT","ThirdTier":"Monitor","OtherDetails":"Group effort with Biomed as lead - refer to checklist."},{"Title":"EKG Carts","SpecificScenario":"Data not transfer from/to carts and CardioServer","FirstTier":"IT","SecondTier":"Biomed","ThirdTier":"Vendors (Philips and/or Epiphany)","OtherDetails":""},{"Title":"EKG Carts","SpecificScenario":"Printing problem from cart","FirstTier":"Biomed","SecondTier":"Philips","ThirdTier":"","OtherDetails":""},{"Title":"EKG Carts","SpecificScenario":"Not connecting to Wifi or Ethernet network","FirstTier":"Biomed","SecondTier":"IT","ThirdTier":"Philips","OtherDetails":""},{"Title":"EKG Carts","SpecificScenario":"New cart setup","FirstTier":"Biomed","SecondTier":"IT","ThirdTier":"Philips","OtherDetails":"Biomed leads this and check with IT to test"},{"Title":"EKG Carts","SpecificScenario":"Bar code scanner","FirstTier":"Biomed","SecondTier":"Philips","ThirdTier":"IT","OtherDetails":""},{"Title":"ID Badge System","SpecificScenario":"PC crashes - needs recovery","FirstTier":"IT","SecondTier":"","ThirdTier":"","OtherDetails":"REinstall windows 10"},{"Title":"ID Badge System","SpecificScenario":"Problem with wack-ass thingy S2 Netbox system","FirstTier":"FM","SecondTier":"JPW","ThirdTier":"IT","OtherDetails":""},{"Title":"ID Badge System","SpecificScenario":"Badge Printer troubleshooting PC connection","FirstTier":"IT","SecondTier":"FM","ThirdTier":"JPW","OtherDetails":""},{"Title":"ID Badge System","SpecificScenario":"Badget Printer equipment issue","FirstTier":"FM","SecondTier":"JPW","ThirdTier":"","OtherDetails":""},{"Title":"ID Badge System","SpecificScenario":"Physical system not working (i.e. door access not working)","FirstTier":"FM","SecondTier":"JPW","ThirdTier":"","OtherDetails":""},{"Title":"ID Badge System","SpecificScenario":"S2 / Parking gate issues","FirstTier":"FM","SecondTier":"JPW","ThirdTier":"","OtherDetails":""},{"Title":"ID Badge System","SpecificScenario":"Security door swipe access programming","FirstTier":"FM","SecondTier":"JPW","ThirdTier":"","OtherDetails":""},{"Title":"ID Badge System","SpecificScenario":"Changing access/lock schedules","FirstTier":"FM (Doug)","SecondTier":"JPW","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Rebooting departmental camera station after power blip or other reason","FirstTier":"FM","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Rebooting physical camera server","FirstTier":"IT (only if KR asks - check first)","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Network related issues- all workstations cannot connect to server","FirstTier":"IT","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Network related issues- singular camera not connecting to server","FirstTier":"IT","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Broken / non-functioning camera","FirstTier":"FM","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Server Windows updates","FirstTier":"KR","SecondTier":"","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Departmental station windows update - not on internet - not required","FirstTier":"N/A","SecondTier":"","ThirdTier":"","OtherDetails":""},{"Title":"Security Cameras","SpecificScenario":"Server Backups","FirstTier":"John S investigating","SecondTier":"","ThirdTier":"","OtherDetails":""},{"Title":"ELPAS(patient wandering)","SpecificScenario":"Bracelet assignment to patient","FirstTier":"Dept Staff","SecondTier":"","ThirdTier":"","OtherDetails":""},{"Title":"ELPAS(patient wandering)","SpecificScenario":"Bracelet issues","FirstTier":"Biomed","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"ELPAS(patient wandering)","SpecificScenario":"Departmental ELPAS computer problem","FirstTier":"FM","SecondTier":"Biomed","ThirdTier":"KR","OtherDetails":""},{"Title":"ELPAS(patient wandering)","SpecificScenario":"Server reboot (virtual in IT data centre)","FirstTier":"IT (only if KR asks - check first)","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"ELPAS(patient wandering)","SpecificScenario":"Server Backups","FirstTier":"John S investigating","SecondTier":"","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call Stratford","SpecificScenario":"RESPONDER System @ SGH Call bells not working (i.e. no alerts, no lights, no sound, cords, domes, buttons,nursing stn phone)","FirstTier":"Biomed","SecondTier":"Gordon Ruth","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call Stratford","SpecificScenario":"No alert at communications station","FirstTier":"Biomed","SecondTier":"Gordon Ruth","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call Stratford","SpecificScenario":"Call bells not going to Wifi Phones","FirstTier":"IT","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call Stratford","SpecificScenario":"Issues with call escalation","FirstTier":"IT","SecondTier":"Connexall","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call CPH,STM,SCH","SpecificScenario":"AUSTO System @ CPH,STM, SCH Call bells not working (i.e. no alerts, no lights, no sound, cords, domes, buttons, nursing stn phone)","FirstTier":"FM","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call CPH,STM,SCH","SpecificScenario":"No alert at communications station","FirstTier":"FM","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call CPH,STM,SCH","SpecificScenario":"Issues with call escalation","FirstTier":"IT","SecondTier":"Connexall","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call CPH,STM,SCH","SpecificScenario":"Marquee boards at switchboard","FirstTier":"FM","SecondTier":"KR","ThirdTier":"","OtherDetails":""},{"Title":"Nurse Call CPH,STM,SCH","SpecificScenario":"Call bells not going to Wifi Phones","FirstTier":"IT","SecondTier":"KR","ThirdTier":"","OtherDetails":""}]';
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
