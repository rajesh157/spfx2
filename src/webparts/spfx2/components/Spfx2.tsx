import * as React from 'react';
import styles from './Spfx2.module.scss';
import { ISpfx2Props } from './ISpfx2Props';
import {ISpfxState} from './ISpfx2State';
import { escape } from '@microsoft/sp-lodash-subset';
import { IButtonProps, DefaultButton } from 'office-ui-fabric-react/lib/Button';   
import { sp } from '@pnp/sp';    
import { getGUID } from "@pnp/common";  
import { autobind } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 }
};


export default class Spfx2 extends React.Component<ISpfx2Props, ISpfxState> {
  constructor(props: ISpfx2Props, state: ISpfxState){
    super(props);
    this.state = {
      addUsers: [],
      options: [],
      category: ""
    };
  }
  public componentDidMount() {  
     /* const options = [
      { key: 'fruitsHeader', text: 'Fruits', itemType: DropdownMenuItemType.Header },
      { key: 'apple', text: 'Apples' },
      { key: 'banana', text: 'Banana' },
      { key: 'orange', text: 'Orange', disabled: true },
      { key: 'grape', text: 'Grape' },
      { key: 'divider_1', text: '-', itemType: DropdownMenuItemType.Divider },
      { key: 'vegetablesHeader', text: 'Vegetables', itemType: DropdownMenuItemType.Header },
      { key: 'broccoli', text: 'Broccoli' },
      { key: 'carrot', text: 'Carrot' },
      { key: 'lettuce', text: 'Lettuce' }
    ];*/
    var vArr =[];
    vArr.length = 0;
    sp.web.lists.getByTitle("Category").items.get().then((items: any[]) => {
      items.forEach((item: any)=>vArr.push({key: item.Id, text: item.Title}));
  });
  
    


    this.setState({options: vArr});
  }
  public render(): React.ReactElement<ISpfx2Props> {
    return (
      <div className={ styles.spfx2 }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <PeoplePicker    
                context={this.props.context}    
                titleText="People Picker"    
                personSelectionLimit={1}    
                groupName={""} // Leave this blank in case you want to filter from all users    
                showtooltip={true}    
                isRequired={true}    
                disabled={false}    
                ensureUser={true}    
                selectedItems={this._getPeoplePickerItems}    
                showHiddenInUI={false}    
                principalTypes={[PrincipalType.User]}    
                resolveDelay={1000} />   
                <Dropdown placeholder="Select an option" onChanged={this.ddlCatChanged} label="Basic uncontrolled example" options={this.state.options} styles={dropdownStyles} />
              <DefaultButton    
                data-automation-id="addSelectedUsers"    
                title="Add Selected Users"    
                onClick={this.addSelectedUsers}>    
                Add Selected Users    
              </DefaultButton>
              {this.state.addUsers}
            </div>
          </div>
        </div>
      </div>
    );
  }

  @autobind   
private ddlCatChanged(option: IDropdownOption, index?: number): void { 
  var strCat: any = option.key;
  this.setState({category : strCat});
}
  @autobind   
private addSelectedUsers(): void {   
  sp.web.lists.getByTitle("Employees").items.add({  
    Title: getGUID(),
    TeamsId: {
      results: this.state.addUsers
  },
  CategoryId: this.state.category,
  MangerId: 3
  }).then(i => {  
      alert(i);
  }).catch(e => { alert(e); });  



}
@autobind
private _getPeoplePickerItems(items: any[]){


 for (let item in items)
  {   
    alert(items[item].id);
    var myArr = this.state.addUsers.slice();
    myArr.push(items[item].id);
    //this.state.addUsers.push(items[item].id);
    this.setState({addUsers:myArr});
    //this.setState({addUsers:items[item].id});
    
  }
}
}
