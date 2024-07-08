import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import FormCustomizer, { IFormCustomizerProps } from './components/FormCustomizer';


export interface IFormCustomizerFormCustomizerProperties {
  sampleText?: string;
}

const LOG_SOURCE: string = 'FormCustomizerFormCustomizer';

export default class FormCustomizerFormCustomizer extends BaseFormCustomizer<IFormCustomizerFormCustomizerProperties> {
  private sp: SPFI;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Activated FormCustomizerFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    this.sp = spfi().using(SPFx({pageContext: this.context.pageContext}));
    return Promise.resolve();
  }

  public render(): void {
    const formCustomizer: React.ReactElement<{}> = React.createElement(FormCustomizer, {
      sp: this.sp,
      context: this.context,
      displayMode: this.displayMode,
      listGuid: this.context.list.guid,
      itemID: this.context.itemId,
      onSave: this._onSave,
      onClose: this._onClose
    } as unknown as IFormCustomizerProps);

    ReactDOM.render(formCustomizer, this.domElement);
  }

  public onDispose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {
    this.formSaved();
  }

  private _onClose = (): void => {
    this.formClosed();
  }
}
