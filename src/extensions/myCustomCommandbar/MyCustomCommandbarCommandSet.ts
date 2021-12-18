import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { SPComponentLoader } from '@microsoft/sp-loader';


import * as strings from 'MyCustomCommandbarCommandSetStrings';



/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMyCustomCommandbarCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'MyCustomCommandbarCommandSet';

export default class MyCustomCommandbarCommandSet extends BaseListViewCommandSet<IMyCustomCommandbarCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MyCustomCommandbarCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {

    const compareOneCommand: Command = this.tryGetCommand('CopyPath');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'CopyPath':
        var itemRelative = event.selectedRows[0].getValueByName("FileRef");
        itemRelative = itemRelative.substring(0,itemRelative.lastIndexOf('/',itemRelative.lastIndexOf('/',itemRelative.lastIndexOf('/')-1)-1));
        //Dialog.alert(`${itemRelative}`);
        var url = this.context.pageContext.web.absoluteUrl.replace(itemRelative,"") + event.selectedRows[0].getValueByName("FileRef");
        navigator.clipboard.writeText(url);
        this.showToastr();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private async showToastr() {
    let toastr: any = await import(
      'toastr'
    );
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css');
    toastr.info("Path copied. Paste with Ctrl+V.");
  }
}
