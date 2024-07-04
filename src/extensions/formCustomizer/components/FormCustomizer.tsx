import * as React from 'react';
import { FormDisplayMode,Guid} from '@microsoft/sp-core-library';
// import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import styles from './FormCustomizer.module.scss';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import NewForm from '../Forms/NewForm';
import EditForm from '../Forms/EditForm';
import DisplayForm from '../Forms/DisplayForm';

export interface IFormCustomizerProps {
  sp:SPFI
  // context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  listGuid: Guid;
  itemID: number;
  onSave: () => void;
  onClose: () => void;
}



const FormCustomizer: React.FC<IFormCustomizerProps> = (props) => {
  

  return (<div className={styles.formCustomizer}>
    {props.displayMode === FormDisplayMode.New &&
        <NewForm sp={props.sp} listGuid={props.listGuid} onSave={props.onSave}
            onClose={props.onClose} />
    }
    {props.displayMode === FormDisplayMode.Edit &&
        <EditForm sp={props.sp} listGuid={props.listGuid} itemId={props.itemID}
            onSave={props.onSave} onClose={props.onClose} />
    }
    {props.displayMode === FormDisplayMode.Display &&
        <DisplayForm sp={props.sp} listGuid={props.listGuid} itemId={props.itemID}
            onClose={props.onClose} />
    }
</div>);
};

export default FormCustomizer
