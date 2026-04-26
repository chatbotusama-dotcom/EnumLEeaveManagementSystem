import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp/presets/all';
import LeaveManagementSystem from './components/LeaveManagementSystem';
export default class LeaveManagementSystemWebPart
  extends BaseClientSideWebPart<{}> {

  public render(): void {
    const element: React.ReactElement = React.createElement(
      LeaveManagementSystem,
      {}
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    sp.setup({
      spfxContext: this.context as any
    });
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}