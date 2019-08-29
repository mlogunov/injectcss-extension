import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'InjectCssApplicationCustomizerStrings';

const LOG_SOURCE: string = 'InjectCssApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IInjectCssApplicationCustomizerProperties {
  // This is an example; replace with your own property
  cssFileUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class InjectCssApplicationCustomizer
  extends BaseApplicationCustomizer<IInjectCssApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    const cssFileUrl: string = this.properties.cssFileUrl;
    if(cssFileUrl){
      const head: HTMLElement = document.getElementsByTagName('head')[0] || document.documentElement;
      let customCssLink: HTMLLinkElement = document.createElement('link');
      customCssLink.href = cssFileUrl;
      customCssLink.rel = 'stylesheet';
      customCssLink.type = 'text/css';
      head.insertAdjacentElement('beforeend', customCssLink);
    }

    return Promise.resolve();
  }
}
