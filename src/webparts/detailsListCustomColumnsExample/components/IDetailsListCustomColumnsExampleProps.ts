import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ClientMode } from './ClientMode';

export interface IDetailsListCustomColumnsExampleProps {
  clientMode: ClientMode;
  context: WebPartContext;
}
