import * as React from 'react';
import styles from './Blocket.module.scss';
import { IBlocketProps } from './IBlocketProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { setup as pnpSetup } from '@pnp/common';
import { sp, Item, Items } from '@pnp/sp';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { Panel } from 'office-ui-fabric-react/lib/Panel';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { string } from 'prop-types';
import { DefaultButton, PrimaryButton, DatePicker } from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';

import { taxonomy, ITermStore, ITermSet, ITermSetData } from "@pnp/sp-taxonomy";
import { CurrentUser } from '@pnp/sp/src/siteusers';

export interface IBlocketState {
  items: any;
  Title: string;
  Description: string;
  Price: string;
  Advertiser: string;
  AdvertiserId: string;
  Categorys: string;
  Search: boolean;
  Sort: boolean;
  SearchItem: any;
  showPanel: boolean;
  newForm: boolean;
  ItemId: string;
  CurrentUser: string;
  ItemAdvertiserId: string;
  IsAdmin: boolean;
}

// termstore
// const store: ITermStore = taxonomy.termStores.getByName("Taxonomy_AhKWcV1J9XJgYtf/sZBNnQ==");
// const set: ITermSet = store.getTermSetById("e4099ad45-2387-47e6-8148-8e7f69638aa2");
// const setWithData: ITermSet & ITermSetData = await set.get();

export default class Blocket extends React.Component<IBlocketProps, IBlocketState> {
  private options: IDropdownOption[];
  constructor(props: IBlocketProps, state: IBlocketState) {
    super(props);

    this.state = {
      items: [],
      Title: '',
      Description: '',
      Price: '',
      Advertiser: '',
      AdvertiserId: '',
      Categorys: '',
      Search: false,
      Sort: false,
      SearchItem: [],
      showPanel: false,
      newForm: false,
      ItemId: '',
      CurrentUser: '',
      ItemAdvertiserId: '',
      IsAdmin: null,
    };
    this.options = [
      { key: 'Fordon', text: 'Fordon' },
      { key: 'Mat', text: 'Mat' },
      { key: 'Människor', text: 'Människor' },
    ];
  }

  public componentDidMount() {
    this.getItems();
  }

  private _showPanel = (item: any): void => {
    //console.log(item);
    this.setState({
      showPanel: true,
      newForm: false,
      Title: item.Title,
      Description: item.Description,
      Price: item.Price,
      Advertiser: item.Advertiser.Title,
      ItemId: item.Id,
      ItemAdvertiserId: item.AdvertiserId,
    });
  }

  private _hidePanel = (): void => {
    this.setState({
      showPanel: false,
      Title: '',
      Description: '',
      Price: '',
      Advertiser: ''
    });
  }

  public addAnnons = (e: any): void => {
    e.preventDefault();

    sp.web.lists.getByTitle("Annonser").items.add({
      Title: this.state.Title,
      Description: this.state.Description,
      Price: this.state.Price,
      Categorys: this.state.Categorys,
      AdvertiserId: this.state.AdvertiserId,
    }).then(() => {
      this.getItems();
      this._hidePanel();
    });
    alert("Advertisment with title: " + document.getElementById('Title')["value"] + " Created !");
  }

  public uppdateAnnons = (id): void => {
    console.log(this.state);
    sp.web.lists.getByTitle('Annonser').items.getById(id).update({
      Title: this.state.Title,
      Description: this.state.Description,
      Price: this.state.Price,
      Categorys: this.state.Categorys,
      AdvertiserId: this.state.AdvertiserId,
    }).then(() => {
      this.getItems();
      this._hidePanel();
    });
  }

  public deleteAnnons = (id): void => {
    console.log(id);
    sp.web.lists.getByTitle("Annonser").items.getById(id).delete()
      .then(() => {
        this._hidePanel();
        this.getItems();
      });
    alert("Annons med title: " + document.getElementById('Title')["value"] + " Deleted !");
  }

  private sortMethod = (): void => {
    if (this.state.Sort === true) {
      this.setState({
        SearchItem: this.state.SearchItem.sort((a, b) => {
          return (a.Price - b.Price)
        }),
        Sort: false
      });
    } else {
      this.setState({
        SearchItem: this.state.SearchItem.sort((a, b) => {
          return (b.Price - a.Price)
        }),
        Sort: true
      });
    }
  }

  private searchAnnons = async (e: any) => {
    // e.preventDefault();
   
    let foundItems = await this.state.items.filter(item => {
      return (
        item.Title.toLowerCase().includes(e.toLowerCase())) ||
        item.Categorys.toLowerCase().includes(e.toLowerCase())
    });

    foundItems !== [] ? this.setState({ SearchItem: foundItems, Search: true }) : null;
  }

  private getItems = (): void => {
    sp.web.lists.getByTitle('Annonser').items.select('*', 'Advertiser/Title', 'Advertiser/Id').expand('Advertiser').get()
      .then((result: any) => {
        console.log(result);
        this.setState({
          items: result
        });
      });
    sp.web.currentUser.get()
      .then((user: CurrentUser) => {
        //console.log(user);
        this.setState({
          AdvertiserId: user['Id'],
          IsAdmin: user['IsSiteAdmin']
        });
      });
  }

  private getPeoplePickerItems = (items: any[]) => {
    console.log(items[0].id);
    this.setState({
      AdvertiserId: items[0].id
    });
  }

  private handlechange = (e: React.FormEvent<HTMLInputElement>, newValue?: string) => {
    console.log(newValue);
    this.setState(prevState => ({
      ...prevState,
      [e.target['id']]: newValue
    }));
  }
  private _dropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    this.setState({
      Categorys: option.text
    });
  };

  public render(): React.ReactElement<IBlocketProps> {

    let values: JSX.Element = undefined;
    if (this.state.Search === true) {
      values = this.state.SearchItem.map(item => {
        return (
          <tr onClick={() => { this._showPanel(item) }}>
            <td > {item.Title}</td>
            <td > {item.Categorys}</td>
            <td > {item.Advertiser.Title}</td>
            <td > {item.Price} Kr</td>
            {/* <td > {item.Description}</td> */}
            <td >{item.Created.slice(0, 10)}  {item.Created.substring(11, 16)}</td>
          </tr>
        );
      });
    } else {
      values = this.state.items.map(item => {
        return (
          <tr onClick={() => { this._showPanel(item) }}>
            <td > {item.Title}</td>
            <td > {item.Categorys}</td>
            <td > {item.Advertiser.Title}</td>
            <td > {item.Price} Kr</td>
            {/* <td > {item.Description}</td> */}
            <td >{item.Created.slice(0, 10)}  {item.Created.substring(11, 16)}</td>
          </tr>
        );
      });
    }

    let formButton: JSX.Element = !this.state.newForm ? (
      <div>
        <PrimaryButton text="Uppdate" type="submit" onClick={() => { this.uppdateAnnons(this.state.ItemId) }} />
        <PrimaryButton text="Delete" type="submit" onClick={() => { this.deleteAnnons(this.state.ItemId) }} />
      </div>
    ) : <PrimaryButton text="Save" type="submit" onClick={this.addAnnons} />

    let showFormButtons = this.state.ItemAdvertiserId === this.state.AdvertiserId || this.state.newForm ? formButton : null

    let isAdmin: JSX.Element = this.state.IsAdmin === true ? (
      <PeoplePicker
        context={this.props.context}
        titleText="Ansvarig"
        personSelectionLimit={1}
        groupName={""}
        showtooltip={true}
        isRequired={true}
        disabled={false}
        defaultSelectedUsers={[this.state.Advertiser]}
        selectedItems={this.getPeoplePickerItems}
        showHiddenInUI={false}
        principalTypes={[PrincipalType.User]}
        ensureUser={true}
        resolveDelay={1000}
      />
    ) : null

    return (
      <div className={styles.blocket}>
        <h3>Bästa Annons Sidan i Sverige</h3>
          <SearchBox
            name="Search"
            placeholder="Search for Title"
            styles={{ root: { width: '100%' } }}
            underlined={true}
            onSearch={(e)=> {this.searchAnnons(e)}}
          />
          {/* <PrimaryButton text="Search" type="submit" onClick={this.searchAnnons} /> */}
          <PrimaryButton text="Sort" type="button" onClick={this.sortMethod} />
      
        <br />
        <PrimaryButton text="Skapa Annons" onClick={() => { this.setState({ showPanel: true, newForm: true }) }} />

        <form>
          <Panel
            isOpen={this.state.showPanel}
            closeButtonAriaLabel="Close"
            isLightDismiss={true}
            headerText="Details"
            onDismiss={this._hidePanel}
          >
            <TextField
              label="Title"
              id="Title"
              value={this.state.Title}
              onChange={this.handlechange}
              styles={{ fieldGroup: { width: 300 } }}
            />
            <TextField
              label="Description"
              id="Description"
              value={this.state.Description}
              onChange={this.handlechange}
              styles={{ fieldGroup: { width: 300 } }}
            />
            <TextField
              label="Price"
              id="Price"
              value={this.state.Price}
              onChange={this.handlechange}
              styles={{ fieldGroup: { width: 300 } }}
            />
            <Dropdown
              placeholder="Categorys"
              label="Status"
              options={this.options}
              onChange={this._dropdownChange}
              defaultSelectedKey={this.state.Categorys}
            />
            {isAdmin}
            {showFormButtons}
          </Panel>
        </form>

        <table className={styles.annons}>
          <thead>
            <tr>
              <th>Title</th>
              <th>Kategori</th>
              <th>Ansvarig</th>
              <th>Pris</th>
              {/* <th>Beskrivning</th> */}
              <th>Datum- tid</th>
            </tr>
          </thead>
          <tbody>
            {values}
          </tbody>
        </table>
      </div>
    );
  }
}
