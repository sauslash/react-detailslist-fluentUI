import * as React from 'react';
import { IHomeDetailsListProps } from './IHomeDetailsListProps';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { sp } from "@pnp/sp";

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});

const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };

export interface IDetailsListBasicExampleItem {
  ID: number;
  name: string;
  category: string;
  amount: number;
}

export interface IDetailsListBasicExampleState {
  items: IDetailsListBasicExampleItem[];
  selectionDetails: string;
}

export default class HomeDetailsList extends React.Component<IHomeDetailsListProps, IDetailsListBasicExampleState> {
  private _selection: Selection;
  private _allItems: IDetailsListBasicExampleItem[];
  private _columns: IColumn[];
  
  constructor(props: IHomeDetailsListProps) {
    super(props);

    this.state = {
      items: [],
      selectionDetails: ''
    };

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });    

    this._columns = [
      { key: 'column1', name: 'ID', fieldName: 'ID', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column3', name: 'Category', fieldName: 'category', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column4', name: 'Amount', fieldName: 'amount', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

  }

  public componentDidMount() {

    sp.web.lists
    .getByTitle("products")
    .items.top(5000)
    .select("ID,name,category,amount")
    .orderBy("name",true)
    .get()
    .then(itens => {
      this._allItems = itens;
      this.setState({
        items: itens,
        selectionDetails: this._getSelectionDetails(),
      });

    })
    .catch(e => {
      console.log("erro", e);
    });

  }

  public render(): React.ReactElement<IHomeDetailsListProps> {
    const { items, selectionDetails } = this.state;

    return (
      <Fabric>
        <div className={exampleChildClass}>{selectionDetails}</div>
        <Announced message={selectionDetails} />
        <TextField
          className={exampleChildClass}
          label="Filter by name:"
          onChange={this._onFilter}
          styles={textFieldStyles}
        />
        <Announced message={`Number of items after filter applied: ${items.length}.`} />
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={items}
            columns={this._columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            onItemInvoked={this._onItemInvoked}
          />
        </MarqueeSelection>
      </Fabric>
    );
  }
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IDetailsListBasicExampleItem).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems,
    });
  }

  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    alert(`Item invoked: ${item.name}`);
  }
}
