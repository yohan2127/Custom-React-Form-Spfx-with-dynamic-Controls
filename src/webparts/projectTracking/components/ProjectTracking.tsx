import * as React from 'react';
//import styles from './ProjectTracking.module.scss'
import { IProjectTrackingProps } from './IProjectTrackingProps';
//import { escape } from '@microsoft/sp-lodash-subset';
//import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPnPPeoplePickerState } from './IPnPPeoplePickerState';
import "@pnp/polyfill-ie11";
import { sp, Web } from '@pnp/sp';
//import { IButtonProps, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import * as $ from 'jquery';
//import { autobind, DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react';
//import { getGUID } from "@pnp/common";
//import { ThemeProvider } from '@microsoft/sp-component-base';
//import { mynewnumber, MYchoices} from '../../../models';


require('./Styles/capstatus.css');
let ProjectMapId:any;

// declare var $:any;
// Import button component      
//var moment = require('moment');
//import * as moment from 'moment';
/*const DayPickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker',

  isRequiredErrorMessage: 'Start date is required.',

  invalidInputErrorMessage: 'Invalid date format.'
};

var listItems="";
var listItems1="";
var listItems2="";
var listItems5="";
*/
var search = window.location.search;
var params = new URLSearchParams(search);
var ProjectnameINd = params.get('ProjectName');
var managerlistid: number;
export default class ProjectTracking extends React.Component<IProjectTrackingProps, IPnPPeoplePickerState> {
  //private _opchoices: MYchoices[]=[];
  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {
      sp.setup({
        sp: {
          headers: {
            "Accept": "application/json; odata=nometadata"
          }
        }
      });
      resolve();
    });
  }
  constructor(props: IProjectTrackingProps, state: IPnPPeoplePickerState) {
    super(props);
    console.log(this.props)
    this.state = {
     // addUsers: [],
      Projectstatus: [{ mymilestone: "", myscore: "", id: "",Forcastedate:"",Actualdate:"",TgtResDate:"" }],
      ProjectIssues: [{ myRisk: "", Mydescp: "", id: "",micplan:"",tgtdate:"",owner:"" }],
     // defaultmyusers: [],
     // firstDayOfWeek: DayOfWeek.Sunday,
    //  value: null,
      overstatus: [{ Choices: "" }],
      DomainStatus: [{ Choices: "" }],
      ScheduleStatus: [{ Choices: "" }],
      RiskStatus: [{ Choices: "" }],
    };
    
    this.fetchdatas = this.fetchdatas.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
    this.updateprojectdetails = this.updateprojectdetails.bind(this);

    this.AddIssues = this.AddIssues.bind(this);
    this.UpdateIssues = this.UpdateIssues.bind(this);

   this.addSelectedUsers = this.addSelectedUsers.bind(this);
   this.updateSelectedUsers = this.updateSelectedUsers.bind(this);
 // this.mydata();
   
  }
 componentDidMount(){

  this._drpdown() ;
 }
  mybutton() {
    if (!ProjectnameINd) {
      return <button
        onClick={this.addSelectedUsers}> 
        Add Project
      </button>
    } else {

      return <button
      onClick={() => this.updateSelectedUsers(ProjectnameINd)}>
        Update Project
    </button>
    }

  }
  createUI() {
    let widthstyle = {
      width:"100%"
    };
    let heightstyle = {
      width:"100%"
    };

    return this.state.Projectstatus.map((el, i) => (
<tr  key={i}>
              <td><input type="text" style={widthstyle} mycustomattribute={el.id || ''} placeholder="MileStone" name="mymilestone" value={el.mymilestone || ''} onChange={this.handleChange.bind(this, i)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={el.id || ''} placeholder="Percentage" name="myscore" value={el.myscore || ''} onChange={this.handleChange.bind(this, i)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={el.id || ''} placeholder="Forcastedate" name="Forcastedate" value={el.Forcastedate || ''} onChange={this.handleChange.bind(this, i)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={el.id || ''} placeholder="Actualdate" name="Actualdate" value={el.Actualdate || ''} onChange={this.handleChange.bind(this, i)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={el.id || ''} placeholder="TgtResDate" name="TgtResDate" value={el.TgtResDate || ''} onChange={this.handleChange.bind(this, i)}/></td>
              <td><span><a href="" mycustomattribute={el.id || ''} onClick={(e) => {
     this.removeClick(e,this, i, el.id)}} >-</a></span></td>
              </tr>
   /*   <div key={i}>
        <input mycustomattribute={el.id || ''} placeholder="MileStone" name="mymilestone" value={el.mymilestone || ''} onChange={this.handleChange.bind(this, i)} />
        <input mycustomattribute={el.id || ''} placeholder="Score" name="myscore" value={el.myscore || ''} onChange={this.handleChange.bind(this, i)} />
        <input mycustomattribute={el.id || ''} type='button' value='remove' onClick={this.removeClick.bind(this, i, el.id)} />
    //  </div> */
    ))
  }
  createIssues() {
    let widthstyle = {
      width:"100%"
    };
    let heightstyle = {
      width:"100%"
    };
//myRisk: "", Mydescp: "", id: "",micplan:"",tgtdate:"",owner:""
    return this.state.ProjectIssues.map((e2, j) => (
<tr  key={j}>
              <td><input type="text" style={widthstyle} mycustomattribute={e2.id || ''} placeholder="Risk/Issues" name="myRisk" value={e2.myRisk || ''} onChange={this.handleIssueChange.bind(this, j)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={e2.id || ''} placeholder="Description" name="Mydescp" value={e2.Mydescp || ''} onChange={this.handleIssueChange.bind(this, j)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={e2.id || ''} placeholder="Owner" name="owner" value={e2.owner || ''} onChange={this.handleIssueChange.bind(this, j)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={e2.id || ''} placeholder="Migration plan" name="micplan" value={e2.micplan || ''} onChange={this.handleIssueChange.bind(this, j)}/></td>
              <td><input type="text" style={widthstyle} mycustomattribute={e2.id || ''} placeholder="TargetDate" name="tgtdate" value={e2.tgtdate || ''} onChange={this.handleIssueChange.bind(this, j)}/></td>
              <td><span><a href="" mycustomattribute={e2.id || ''} onClick={(e) => {
     this.removeIssueClick(e,this, j, e2.id)}} >-</a></span></td>
              </tr>
    ))
  }

  removeIssueClick(event,mythis,j, delid) {
    event.preventDefault();
    
    if (ProjectnameINd) {
      let list = sp.web.lists.getByTitle("ProjectIssues");

      list.items.getById(delid).delete().then(_ => {

        console.log("Deleted")
      });
    }
    let ProjectIssues = [...this.state.ProjectIssues];
    ProjectIssues.splice(j, 1);
    this.setState({ ProjectIssues });
  }

  handleChange(i, e) {
    const { name, value } = e.target;
    let Projectstatus = [...this.state.Projectstatus];
    Projectstatus[i] = { ...Projectstatus[i], [name]: value };
    this.setState({ Projectstatus });
  }

  handleIssueChange(i, e) {
    const { name, value } = e.target;
    let ProjectIssues = [...this.state.ProjectIssues];
    ProjectIssues[i] = { ...ProjectIssues[i], [name]: value };
    this.setState({ ProjectIssues });
  }
  addClick() {
    this.setState(prevState => ({
      Projectstatus: [...prevState.Projectstatus, { mymilestone: "", myscore: "", id: "",Forcastedate:"",Actualdate:"",TgtResDate:"" }]
    }))
  }
  addIssueClick() {
    this.setState(prevState => ({
      ProjectIssues: [...prevState.ProjectIssues, { myRisk: "", Mydescp: "", id: "",micplan:"",tgtdate:"",owner:"" }]
    }))
  }
  private    _drpdown()   {
   
      const web1 = new Web(this.props.context.pageContext.web.absoluteUrl);
      let batch = web1.createBatch();  
      web1.lists.getByTitle("Program").fields.getByInternalNameOrTitle("RiskStatus").select('Choices')
      .inBatch(batch).get().then((fieldData5) => {
       this.setState({ RiskStatus: fieldData5 });
      });
      web1.lists.getByTitle("Program").fields.getByInternalNameOrTitle("OverAllStatus").select('Choices')
    .inBatch(batch).get().then((fieldData4) => {
        this.setState({ overstatus: fieldData4 });
      });
      web1.lists.getByTitle("Program").fields.getByInternalNameOrTitle("DomainStatus").select('Choices')
      .inBatch(batch).get().then((fieldData3) => {
        this.setState({ DomainStatus: fieldData3 });
      });
      web1.lists.getByTitle("Program").fields.getByInternalNameOrTitle("ScheduleStatus").select('Choices')
      .inBatch(batch).get().then((fieldData2) => {
        this.setState({ ScheduleStatus: fieldData2 });
        
      });
 
      batch.execute().then(() => {
           if (ProjectnameINd)
        this.fetchdatas();
      });
      //this.render();
    }
  
  removeClick(event,mythis,i, delid) {
    event.preventDefault();
    // var delid=Number($(this).attr("ProjectStatus"));
    if (ProjectnameINd) {
      let list = sp.web.lists.getByTitle("ProjectStatus");

      list.items.getById(delid).delete().then(_ => {

        console.log("Deleted")
      });
    }
    let Projectstatus = [...this.state.Projectstatus];
    Projectstatus.splice(i, 1);
    this.setState({ Projectstatus });
  }

  async handleSubmit(myid:number) {
    // alert('A name was submitted: ' + JSON.stringify(this.state.Projectstatus));
    //console.log(this.state.Projectstatus);
    var mystatecopy = [];
    mystatecopy.push(this.state.Projectstatus);
    //console.log(mystatecopy);
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    const batch = web.createBatch();

    const list = web.lists.getByTitle("ProjectStatus");
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    for (let k = 0; k < mystatecopy[0].length; k++) {

      list.items.inBatch(batch).add({
        ProjectNameId: myid,
        MileStone: mystatecopy[0][k].mymilestone,
        PercentageComplete: mystatecopy[0][k].myscore,
        ForcastDate: mystatecopy[0][k].Forcastedate,
        ActualDate: mystatecopy[0][k].Actualdate,
        TargetResolutionDate: mystatecopy[0][k].TgtResDate
      }, entityTypeFullName).then(b => {
        // console.log(b);
      });

    }
    batch.execute().then(() => {
     this.AddIssues(myid);

    });


    //event.preventDefault();
  }

  async AddIssues(myid:number) {
    // alert('A name was submitted: ' + JSON.stringify(this.state.Projectstatus));
    //console.log(this.state.Projectstatus);
    let mystatecopy = [];
    mystatecopy.push(this.state.ProjectIssues);
    //console.log(mystatecopy);
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    const batch = web.createBatch();
    //{ myRisk: "", Mydescp: "", id: "",micplan:"",tgtdate:"",owner:"" }
    const list = web.lists.getByTitle("ProjectIssues");
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    for (let k = 0; k < mystatecopy[0].length; k++) {

      list.items.inBatch(batch).add({
        ProjectNameId: myid,
        Risk: mystatecopy[0][k].myRisk,
        Description: mystatecopy[0][k].Mydescp,
        MigrationPLan: mystatecopy[0][k].micplan,
        Owner: mystatecopy[0][k].owner,
        TargetResolutionDate: mystatecopy[0][k].tgtdate
      }, entityTypeFullName).then(b => {
        // console.log(b);
      });

    }
    batch.execute().then(() => console.log("All done!"));


 
  }

  async UpdateIssues(managerlistid) {
    // alert('A name was submitted: ' + JSON.stringify(this.state.Projectstatus));
    //console.log(this.state.Projectstatus);
    var mystatecopy = [];
    mystatecopy.push(this.state.ProjectIssues);
    //console.log(mystatecopy);
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    const batch = web.createBatch();

    const list = web.lists.getByTitle("ProjectIssues");
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    for (let k = 0; k < mystatecopy[0].length; k++) {
      if (!mystatecopy[0][k].id) {
        list.items.inBatch(batch).add({
          ProjectNameId: managerlistid,
        Risk: mystatecopy[0][k].myRisk,
        Description: mystatecopy[0][k].Mydescp,
        MigrationPLan: mystatecopy[0][k].micplan,
        Owner: mystatecopy[0][k].owner,
        TargetResolutionDate: mystatecopy[0][k].tgtdate
        }, entityTypeFullName).then(b => {
          // console.log(b);
        });
      }
      else {
        list.items.inBatch(batch).getById(mystatecopy[0][k].id).update({
          ProjectNameId: managerlistid,
        Risk: mystatecopy[0][k].myRisk,
        Description: mystatecopy[0][k].Mydescp,
        MigrationPLan: mystatecopy[0][k].micplan,
        Owner: mystatecopy[0][k].owner,
        TargetResolutionDate: mystatecopy[0][k].tgtdate
        }).then(b => {
          // console.log(b);
        });

      }
    }
    batch.execute().then(() => console.log("All done!"));


    // event.preventDefault();
  }

  async updateprojectdetails(managerlistid) {
    // alert('A name was submitted: ' + JSON.stringify(this.state.Projectstatus));
    //console.log(this.state.Projectstatus);
    var mystatecopy = [];
    mystatecopy.push(this.state.Projectstatus);
    //console.log(mystatecopy);
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    const batch = web.createBatch();

    const list = web.lists.getByTitle("ProjectStatus");
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    for (let k = 0; k < mystatecopy[0].length; k++) {
      if (!mystatecopy[0][k].id) {
        list.items.inBatch(batch).add({
          ProjectNameId: managerlistid,
          MileStone: mystatecopy[0][k].mymilestone,
          PercentageComplete: mystatecopy[0][k].myscore,
          ForcastDate: mystatecopy[0][k].Forcastedate,
          ActualDate: mystatecopy[0][k].Actualdate,
          TargetResolutionDate: mystatecopy[0][k].TgtResDate
        }, entityTypeFullName).then(b => {
          // console.log(b);
        });
      }
      else {
        list.items.inBatch(batch).getById(mystatecopy[0][k].id).update({
          ProjectNameId:managerlistid,
          MileStone: mystatecopy[0][k].mymilestone,
          SPercentageComplete: mystatecopy[0][k].myscore,
          ForcastDate: mystatecopy[0][k].Forcastedate,
          ActualDate: mystatecopy[0][k].Actualdate,
          TargetResolutionDate: mystatecopy[0][k].TgtResDate
        }).then(b => {
          // console.log(b);
        });

      }
    }
    batch.execute().then(() => console.log("All done!"));
    this.UpdateIssues(managerlistid);

    // event.preventDefault();
  }


   fetchdatas() {
 
    var reg1 = new RegExp('<div class=\"ExternalClass[0-9A-F]+\">', "");
    var reg2 = new RegExp('</div>$', "");
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    //const batch1 = web.createBatch();

    const list1 = web.lists.getByTitle("Program");
    const list2 = web.lists.getByTitle("ProjectStatus");
    const list3 = web.lists.getByTitle("ProjectIssues");

 let FetchProjectDetails = [];
 let FetchProjectIssues = [];
      list1.items.select('Id,OPSTeam,Charter,RoadMap,OPMTeam,Others,DomainStatus,RiskStatus,ScheduleStatus,ProgramName,Sponsor,OverAllStatus,ProgramManager,ProjectScope,Highlights,Lowlights,ExecutiveStatus,TargetGoLive').filter("ProgramName eq '" + ProjectnameINd + "'").get().then(r => {
      console.log(r)
      let Pjtscope="";
      let HighLghts="";
      let lowlghts="";
      let Extstatus="";
      let Pgmmngr="";
      let opstm="";
      let opmtm="";
      let othrs="";
      let spnsr="";
      if(r[0].ProjectScope)
                    Pjtscope=r[0].ProjectScope.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].Highlights)
                    HighLghts=r[0].Highlights.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].Lowlights)
                    lowlghts=r[0].Lowlights.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].ExecutiveStatus)
                    Extstatus=r[0].ExecutiveStatus.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].ProgramManager)
                    Pgmmngr=r[0].ProgramManager.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].OPSTeam)
                    opstm=r[0].OPSTeam.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].OPMTeam)
                    opmtm=r[0].OPMTeam.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].Others)
                    othrs=r[0].Others.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
                    if(r[0].Sponsor)
                    spnsr=r[0].Sponsor.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
 	
                    
      managerlistid = r[0].Id;
      $("#ProjectName").val(r[0].ProgramName);
      $("#ProjectDetails").val(Pjtscope);
       $("#Risk").val(r[0].RiskStatus);
       $("#Schedule").val(r[0].ScheduleStatus);
       $("#Overall").val(r[0].OverAllStatus);
       $("#Domain").val(r[0].DomainStatus);
       $("#Sponsor").val(spnsr);
       $("#Others").val(othrs);
       $("#OPMTeamId").val(opmtm);
       $("#OPSTeamId").val(opstm);
       $("#Target_go-Live").val(r[0].TargetGoLive);
       $("#Executive").val(Extstatus);
       $("#Lowlights").val(lowlghts);
       $("#Highlights").val(HighLghts);
       $("#ProjectManagerId").val(Pgmmngr); 
       if(r[0].Charter)  
       $("#Charter").val(r[0].Charter.Url);  
       if(r[0].RoadMap) 
       $("#RoadMap").val(r[0].RoadMap.Url);
    });
  
     list2.items.select('Id,MileStone,PercentageComplete,ForcastDate,ActualDate,TargetResolutionDate').filter("ProjectName/ProgramName eq '" + ProjectnameINd + "'").top(5000).get().then(r => {
      for (let i = 0; i < r.length; i++) {
        FetchProjectDetails.push({
          mymilestone: r[i].MileStone,
          myscore: r[i].PercentageComplete,
          id: r[i].Id,
          Forcastedate:r[i].ForcastDate,
          Actualdate:r[i].ActualDate,
          TgtResDate:r[i].TargetResolutionDate

        })
      }
      console.log(r)
      this.setState({ Projectstatus: FetchProjectDetails });

    });
 let ownn="";
 let Tgt="";
    list3.items.select('Id,Risk,Description,MigrationPLan,Owner,TargetResolutionDate').filter("ProjectName/ProgramName eq '" + ProjectnameINd + "'").top(5000).get().then(r => {
     
      for (let i = 0; i < r.length; i++) {
        if(r[i].Owner)
        ownn=r[i].Owner.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');

        if(r[i].TargetResolutionDate)
        Tgt=r[i].TargetResolutionDate.replace(reg1, "").replace(reg2, "").replace(/<(?:.|\n)*?>/gm, '').replace(/[\u200B]/g, '');
        FetchProjectIssues.push({
          myRisk: r[i].Risk,
          Mydescp: r[i].Description,
          id: r[i].Id,
          micplan:r[i].MigrationPLan,
          tgtdate:Tgt,
          owner:ownn

        })
      }
      console.log(r)
      this.setState({ ProjectIssues: FetchProjectIssues });

    });



  }
  public render(): React.ReactElement<IProjectTrackingProps> {
   // const { firstDayOfWeek, value } = this.state;
   /* const renObjData = this.props.mychoices["Choices"].map(function(data, idx) {
      return <p key={idx}>{data}</p>;
  });*/
 /* var  mychoice1="";
  var  mychoice2="";
  var  mychoice3="";
  var  mychoice4="";

  if(this.props.mychoices["Choices"]){
   mychoice1= this.props.mychoices["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); }
    if(this.props.mychoices2["Choices"]){
     mychoice2= this.props.mychoices2["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); }
    if(this.props.mychoices3["Choices"]){
     mychoice3 = this.props.mychoices3["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); }
    if(this.props.mychoices4["Choices"]){
     mychoice4 = this.props.mychoices4["Choices"].map((item, i: number): JSX.Element => {  
      return (  
        <option  value={item || ''}>{item || ''}</option>
       
      );  
    }); } */
    let widthstyle = {
      width:"100%"
    };
    let heightstyle = {
      width:"100%"
    };
    return (
    /* <div>Madhan</div>*/
     <div className="wrapper">
      <div className="container">
        <div className="program_list">
          <div>
            <div><label>Program Manager:</label></div>
            <div><input className="input" type="text" id="ProjectManagerId" /></div>
          </div>
          
          <div>
            <div><label>OPS Team</label></div>
            <div><input className="input" type="text"  maxLength={30}   id="OPSTeamId"/></div>
          </div>
          
          <div>
            <div><label>Target Go-Live</label></div>
            <div><input className="input" type="text"  maxLength={30}  id="Target_go-Live"/></div>
          </div>
          <div>
            <div><label>Sponsor</label></div>
            <div><input className="input" type="text"  maxLength={30}  id="Sponsor"/></div>
          </div>
          <div>
            <div><label>Charter</label></div>
            <div><input className="input" type="text"  maxLength={30}   id="Charter"/></div>
          </div>
          <div id="overalll">
            <div className="overallstatus"><label>Overall Status</label></div>
            <div className="overall_drp">
              <select id="Overall">
              {
                this.state.overstatus["Choices"] && this.state.overstatus["Choices"].map((mynum2) =>
        <option  value={mynum2 || ''}>{mynum2 || ''}</option> )}  
              </select>
            </div>
          </div>
          <div className="Domainn">
            <div className="DomainStatus"><label>Domain Status</label></div>
            <div className="DomainStatus_drp">
              <select id="Domain">
              {this.state.DomainStatus["Choices"] && this.state.DomainStatus["Choices"].map((mynum2) =>
        <option  value={mynum2 || ''}>{mynum2 || ''}</option> )}
              </select>
            </div>
          </div>	
        </div>
        <div className="program">
          <div>
            <div><label>Program Name</label></div>
            <div><input className="input" type="text"  maxLength={30}  id="ProjectName"/></div>
          </div>
          <div>
            <div><label>OPM Team</label></div>
            <div><input className="input" type="text"  maxLength={30}  id="OPMTeamId"/></div>
          </div>
          <div style={heightstyle}>
            <div><label>Executive Status</label></div>
            <div><textarea id="Executive"></textarea></div>
          </div>
          <div>
            <div><label>RoadMap</label></div>
            <div><input id="RoadMap" className="input" type="text"  maxLength={30}  /></div>
          </div>
          <div className="Domainn">
            <div className="DomainStatus"><label>Risk Status</label></div>
            <div className="DomainStatus_drp">
              <select  id="Risk">
              {this.state.RiskStatus["Choices"] && this.state.RiskStatus["Choices"].map((mynum2) =>
       <option  value={mynum2 || ''}>{mynum2 || ''}</option> )} 
              </select>
            </div>
          </div>
          <div className="Domainn">
            <div className="DomainStatus"><label>Schedule Status</label></div>
            <div className="DomainStatus_drp">
              <select id="Schedule">
              {this.state.ScheduleStatus["Choices"] && this.state.ScheduleStatus["Choices"].map((mynum2) =>
        <option  value={mynum2 || ''}>{mynum2 || ''}</option> )}
              </select>
            </div>
          </div>
        </div>
        <div className="prog_scope">
          <div className="pro_scope_title">Project Scope</div>
          <div className="pro_scope_content">
          <input type="text"  id="ProjectDetails"></input>

          </div>
        </div>
        <div className="high_low">
          <div className="hightlight">
            <div className="highlight_title">Highlights</div>
            <div className="highlight_content">

            <input type="text" id="Highlights"></input>
            </div>
          </div>
          <div className="lowlight">
            <div className="lowlight_title">Lowlights</div>
            <div className="lowlight_content">

            <input type="text"  id="Lowlights"></input>
            </div>
          </div>
        </div>
        <div className="milestone">
          <table>
          <tbody>
              <tr>
              <th>Milestone Title</th>
              <th>%Complete</th>
              <th>Forecast Date</th>
              <th>Actual Date</th>
              <th>Target Resolution Date</th>
              </tr>
              <a href="#" id="Addbutton" onClick={this.addClick.bind(this)}>+</a><br></br>
              {this.createUI()}
              </tbody>
          </table>
        </div>
        <div className="risk">
          <table>
            <tbody>
              <tr>
              <th>Risk/Issues</th>
              <th>Description</th>
              <th>Owner</th>
              <th>Mitigation Plan</th>
              <th>Target Resolution Date</th>
              <th></th>
              </tr>
              <a href="#" id="IssueAdd" onClick={this.addIssueClick.bind(this)}>+</a><br></br>
              {this.createIssues()}
              </tbody>
          </table>
        </div>
        {this.mybutton()}     
      
      </div>
    </div>
 );

  }
  private addSelectedUsers(): void{
    
    let projectname = $("#ProjectName").val();
    let projectdetail = $("#ProjectDetails").val();
    let charter :string
    let RoadMap :string
     charter = $("#Charter").val();
     RoadMap = $("#RoadMap").val();
  let chartdesc=charter.split(".");
  var chartdesc1=chartdesc[chartdesc.length-1];
  let RoadMapdesc=RoadMap.split(".");
  var RoadMapdesc1=RoadMapdesc[RoadMapdesc.length-1]
    var AddData ={
      "ProgramName": projectname,
      "ProgramManager":$("#ProjectManagerId").val(),
      //"ProjectDetails":projectdetail,
      "ProjectScope": projectdetail,
      "Highlights": $("#Highlights").val(),
      "Lowlights": $("#Lowlights").val(),
      "ExecutiveStatus": $("#Executive").val(),
      "TargetGoLive":  $("#Target_go-Live").val(),
      Charter: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: chartdesc1,
        Url: charter
      },
      RoadMap: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: RoadMapdesc1,
        Url: RoadMap
      },
     // "ProgramManagerId": { "results": peoplepicarray },
      "OPSTeam": $("#OPSTeamId").val(),
      "OPMTeam": $("#OPMTeamId").val(),
      "Others": $("#Others").val(),
      "Sponsor": $("#Sponsor").val(),
      "OverAllStatus": $("#Overall option:selected").val(),
      "DomainStatus": $("#Domain option:selected").val(),
      "ScheduleStatus": $("#Schedule option:selected").val(),
      "RiskStatus": $("#Risk option:selected").val()
     
      /*  ProjectManager: {   
            results: this.state.addUsers  
        }  */
    };
    sp.web.lists.getByTitle("Program").items.add(AddData).then(i => {
      console.log(i);
      this.handleSubmit(i.data.Id);
    });

  }
  private updateSelectedUsers(ProjectnameINd): void {
   
    let projectname = $("#ProjectName").val();
    let projectdetail = $("#ProjectDetails").val();
    let charter :string
    let RoadMap :string
     charter = $("#Charter").val();
     RoadMap = $("#RoadMap").val();
  let chartdesc=charter.split(".");
  var chartdesc1=chartdesc[chartdesc.length-1];
  let RoadMapdesc=RoadMap.split(".");
  var RoadMapdesc1=RoadMapdesc[RoadMapdesc.length-1]
    var UpdateData ={
      "ProgramName": projectname,
      "ProgramManager":$("#ProjectManagerId").val(),
      //"ProjectDetails":projectdetail,
      "ProjectScope": projectdetail,
      "Highlights": $("#Highlights").val(),
      "Lowlights": $("#Lowlights").val(),
      "ExecutiveStatus": $("#Executive").val(),
      "TargetGoLive":  $("#Target_go-Live").val(),
      Charter: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: chartdesc1,
        Url: charter
      },
      RoadMap: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: RoadMapdesc1,
        Url: RoadMap
      },
     // "ProgramManagerId": { "results": peoplepicarray },
      "OPSTeam": $("#OPSTeamId").val(),
      "OPMTeam": $("#OPMTeamId").val(),
      "Others": $("#Others").val(),
      "Sponsor": $("#Sponsor").val(),
      "OverAllStatus": $("#Overall option:selected").val(),
      "DomainStatus": $("#Domain option:selected").val(),
      "ScheduleStatus": $("#Schedule option:selected").val(),
      "RiskStatus": $("#Risk option:selected").val()
     
      /*  ProjectManager: {   
            results: this.state.addUsers  
        }  */
    };
    sp.web.lists.getByTitle("Program").items.getById(managerlistid).update(UpdateData).then(i => {
      console.log(i);
      this.updateprojectdetails(managerlistid);
    });

  }





  /*private onaddbtns = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  
    this.props.onAddButton();
  }  
  private onUpdatebtns = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
  
    this.props.onDeleteBtn();
  }*/

}
