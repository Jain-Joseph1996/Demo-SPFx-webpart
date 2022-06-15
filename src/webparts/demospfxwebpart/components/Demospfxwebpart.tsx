import * as React from 'react';
import styles from './Demospfxwebpart.module.scss';
import { IDemospfxwebpartProps } from './IDemospfxwebpartProps';
import { IDemospfxwebpartState } from './IDemospfxwebpartState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label, TextField, ChoiceGroup, Checkbox, IChoiceGroupOption, PrimaryButton, Dialog, DialogFooter } from "@fluentui/react";
import { DialogType } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import interfaces
import { IFile, IResponseItem } from "./interfaces";
import { Caching } from "@pnp/queryable";
import { getSP } from "../pnpjsConfig";
import { SPFI, spfi } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import { IItemUpdateResult } from "@pnp/sp/items";


const dialogContent = {
  type: DialogType.normal,
  Title: "Message",
  subText: "Your request susseccfully submitted",
  closeButtonAriaLabel: "Close"
}
export default class Demospfxwebpart extends React.Component<
  IDemospfxwebpartProps,
  IDemospfxwebpartState,
  {}
>
{
  private _sp: SPFI;
  constructor(props: IDemospfxwebpartProps) {
    super(props);
    this.getEmail = this.getEmail.bind(this);
    this.getMobile = this.getMobile.bind(this);
    this.getAddress = this.getAddress.bind(this);
    this.getApproval = this.getApproval.bind(this);
    this.getAvailability = this.getAvailability.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this.submitData = this.submitData.bind(this);
    this.Cancel = this.Cancel.bind(this);
    this.state = {
      email: "",
      mobile: "",
      address: "",
      availability: false,
      approval: "",
      empPeoplepicker: [],
      hidedialog: true,
      defaultemp: [""],
    };
    this._sp = getSP();
  }


  // Arrow function
  public toggleDialog = (ev) => {
    this.setState({
      hidedialog: true
    })
  }

  public choicetype: IChoiceGroupOption[] =
    [{ key: "Yes", text: "Yes" },
    { key: "No", text: "No" }]

  // Read email values
  public getEmail(ev, value: string) {
    this.setState({
      email: value
    })
  }

  // Read mobile values
  public getMobile(ev, value: string) {
    this.setState({
      mobile: value
    })
  }
  // Read address
  public getAddress(ev, value: string) {
    this.setState({
      address: value
    })
  }

  // Read approval field
  public getApproval(ev, value: IChoiceGroupOption) {
    this.setState({
      approval: value.key,
    })
  }

  // Read availability
  public getAvailability(ev, value: boolean) {
    this.setState({
      availability: value
    })
  }

  // get people picker items
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
    let ppl: any[] = [];
    let defaultvalue: string[] = [];
    items.map((item) => {
      ppl.push(item.id);
      defaultvalue.push(item.secondaryText);
    })
    this.setState({
      empPeoplepicker: ppl,
      defaultemp: defaultvalue
    })
  }

  // Submit function
  public async submitData() {
    let valid: boolean = true;
    if (this.state.email == "") {
      valid = false;
      document
        .getElementById("email_valid")
        .setAttribute("style", "display:block!important");
    }
    if (this.state.mobile == "") {
      valid = false;
      document
        .getElementById("mobile_valid")
        .setAttribute("style", "display:block!important");
    }
    if (this.state.empPeoplepicker.length == 0) {
      valid = false;
      document
        .getElementById("employee_valid")
        .setAttribute("style", "display:block!important");
    }

    if (valid) {

      const data = {
        Title: "Custom List Form",
        Employeename: { results: this.state.empPeoplepicker },
        Mobile: this.state.mobile,
        Address: this.state.address,
        Email: this.state.email,
        availability: this.state.availability,
        approval: this.state.approval

      }

      this._sp.web.lists
        .getByTitle("FormDemo").items.add(data)
        .then(() => {
          this.setState({
            hidedialog: false
          })
        })
    }
  }
  public Cancel() {
    this.setState({
      email: "",
      mobile: "",
      address: "",
      availability: false,
      approval: "",
      defaultemp: [],
    });

  }
  public render(): React.ReactElement<IDemospfxwebpartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    return (
      <section className={`${styles.demospfxwebpart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>

          <div>
            <h3>Request Form</h3>
            <div id="custom_form" >
              <div className={styles.grid}>
                <div className={styles.gridRow}>
                  <div className={styles.firstcolumn}><Label>Employee Name<span>*</span></Label></div>
                  <div className={styles.secondColumn}><PeoplePicker
                    context={this.props.context}
                    placeholder="Enter Your Name"
                    ensureUser={true}
                    personSelectionLimit={3}
                    groupName={""}
                    showtooltip={false}
                    disabled={false}
                    showHiddenInUI={false}
                    resolveDelay={1000}
                    principalTypes={[PrincipalType.User]}
                    onChange={this._getPeoplePickerItems}
                    defaultSelectedUsers={this.state.defaultemp}
                  ></PeoplePicker>
                    <div id="employee_valid" className="form-validation"><span>You can't leave this blank</span></div>
                  </div>

                  <div className={styles.firstcolumn}><Label>Email<span>*</span></Label></div>
                  <div className={styles.secondColumn}>
                    <TextField
                      placeholder='Enter email here'
                      onChange={this.getEmail}>
                    </TextField>
                    <div id="email_valid" className="form-validation"><span>You can't leave this blank</span></div>
                  </div>

                  <div className={styles.firstcolumn}><Label>Mobile<span>*</span></Label></div>
                  <div className={styles.secondColumn}>
                    <TextField
                      type='number'
                      placeholder='Enter number here'
                      onChange={this.getMobile}>
                    </TextField>
                    <div id="mobile_valid" className="form-validation"><span>You can't leave this blank</span></div>
                  </div>

                  <div className={styles.firstcolumn}><Label>Address<span>*</span></Label></div>
                  <div className={styles.secondColumn}>
                    <TextField
                      multiline={true}
                      placeholder=''
                      onChange={this.getAddress}>
                    </TextField>
                  </div>

                  <div className={styles.firstcolumn}><Label>Do you have approval?<span>*</span></Label></div>
                  <div className={styles.secondColumn}>
                    <ChoiceGroup
                      options={this.choicetype}
                      onChange={this.getApproval}
                    >
                    </ChoiceGroup>
                    <div id="approval_valid" className="form-validation"><span>You can't leave this blank</span></div>
                  </div>

                  <div className={styles.firstcolumn}><Label>Are you available?<span>*</span></Label></div>
                  <div className={styles.secondColumn}>
                    <Checkbox
                      label="Yes"
                      onChange={this.getAvailability}
                      checked={this.state.availability}>
                    </Checkbox>
                    <div id="available_valid" className="form-validation"><span>You can't leave this blank</span></div>
                  </div>
                  <Dialog
                    dialogContentProps={dialogContent}
                    hidden={this.state.hidedialog}
                    onDismiss={this.toggleDialog}>
                    <DialogFooter>
                      <PrimaryButton
                        text='Close'
                        onClick={this.toggleDialog}>
                      </PrimaryButton>
                    </DialogFooter>
                  </Dialog>
                  <div className={styles.secondColumn}>
                    <PrimaryButton
                      className={styles.button}
                      text='Submit'
                      onClick={this.submitData}>
                    </PrimaryButton>
                    <PrimaryButton
                      className={styles.button}
                      text='Cancel'
                      onClick={this.Cancel}>
                    </PrimaryButton>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section >
    );
  }
}
