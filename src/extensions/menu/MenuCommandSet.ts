import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

// Reference the SP2013 solution
import "jslinkMenu";
declare var JSLinkMenu;

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMenuCommandSetProperties { }

const LOG_SOURCE: string = 'MenuCommandSet';

export default class MenuCommandSet extends BaseListViewCommandSet<IMenuCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MenuCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    let commands = ["COMMAND_1", "COMMAND_2"];

    // Parse the command names
    for (let commandName of commands) {
      // Ensure the command exists
      const command: Command = this.tryGetCommand(commandName);
      if (command) {
        // Display the command if items have been selected
        command.visible = event.selectedRows.length > 0;
      }
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    // Parse the fields
    let fieldNames = [];
    let fields = event.selectedRows.length > 0 ? event.selectedRows[0].fields : [];
    for (let field of fields) {
      // Save the field name
      fieldNames.push(field.internalName);
    }

    // Parse the selected rows
    let items = [];
    for (let selectedItem of event.selectedRows) {
      let item = {};

      // Parse the field names
      for (let fieldName of fieldNames) {
        // Add the item value
        item[fieldName] = selectedItem.getValueByName(fieldName);
      }

      // Add the item
      items.push(item);
    }

    switch (event.itemId) {
      case 'COMMAND_1':
        // Execute command 1
        JSLinkMenu.Commands.Command1(items);
        break;
      case 'COMMAND_2':
        // Execute command 2
        JSLinkMenu.Commands.Command2(items);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
