import * as React from "react";
import { MsalClientSetup } from "@pnp/msaljsclient";
import { IWeb, IList, sp } from "@pnp/sp/presets/all";
import {
  ComboBox,
  IComboBox,
  IComboBoxOption,
  PrimaryButton,
  DefaultButton
} from 'office-ui-fabric-react/lib/index';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { hiddenContentStyle, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { useId } from '@uifabric/react-hooks';
// import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
// import DocsNav from "./DocsNav";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
/* global Button, Header, HeroList, HeroListItem, Progress */


export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  folders: IList;
  listProjects: IComboBoxOption[];
  listFolders: IComboBoxOption[];
  listSubFolders: IComboBoxOption[];
  projectId: number;
  categoryId: number;
  subCategoryId: number;
  btnDisabled: boolean;
  modalHide: boolean;
  mbShow: boolean;
  mbType: MessageBarType;
  mbText: string;
  showSuccess: boolean;
}

export interface ModalProps {
  acceptFunction: () => any;
  cancelFunction: () => any;
  closeFunction: () => any;
  hide: boolean;
}

export interface MessageBarProps {
  closeFunction: () => any;
  choice: MessageBarType;
  text: string;
}

const comboBoxBasicOptions: IComboBoxOption[] = [
  { key: 1, text: 'Option A' },
  { key: 2, text: 'Option B' },
  { key: 3, text: 'Option C' },
  { key: 4, text: 'Option D' },
  { key: 5, text: 'Option E' },
  { key: 6, text: 'Option F', disabled: true },
  { key: 7, text: 'Option G' },
  { key: 8, text: 'Option H' },
  { key: 9, text: 'Option I' },
  { key: 10, text: 'Option J' },
];

const screenReaderOnly = mergeStyles(hiddenContentStyle);
const dialogContentProps = {
  type: DialogType.normal,
  title: 'Missing Subject',
  closeButtonAriaLabel: 'Close',
  subText: 'Do you want to send this message without a subject?',
};

const cbStyle = { width: '75%', display: 'block', margin: '0 auto' };
const btnStyle = { width: '75%', display: 'block', margin: '20px auto 0 auto' };
const dialogStyles = { main: { maxWidth: 450 } };
const successMessageStyles = { width: '80%', heigth: '100%', margin: '0 auto' };
let formContainerStyles = { display: 'block' };


////////////////// MODAL
const DialogBasicExample: React.FunctionComponent<ModalProps> = (props) => {
  const labelId: string = useId('dialogLabel');
  const subTextId: string = useId('subTextLabel');

  const modalProps = React.useMemo(
    () => ({
      titleAriaId: labelId,
      subtitleAriaId: subTextId,
      isBlocking: false,
      styles: dialogStyles
    }),
    [labelId, subTextId],
  );

  return (
    <>
      {/* <DefaultButton secondaryText="Opens the Sample Dialog" onClick={props.closeFunction} text="Open Dialog" /> */}
      <label id={labelId} className={screenReaderOnly}>
        My sample label
      </label>
      <label id={subTextId} className={screenReaderOnly}>
        My sample description
      </label>

      <Dialog
        hidden={props.hide}
        onDismiss={props.closeFunction}
        dialogContentProps={dialogContentProps}
        modalProps={modalProps}
      >
        <DialogFooter>
          <PrimaryButton onClick={props.acceptFunction} text="Confirmar" />
          <DefaultButton onClick={props.cancelFunction} text="Cancelar" />
        </DialogFooter>
      </Dialog>
    </>
  );
};

///////////// MESSAGE BAR

const AddInMessageBar: React.FunctionComponent<MessageBarProps> = (props) => (
  <MessageBar
    messageBarType={props.choice}
    isMultiline={false}
    onDismiss={props.closeFunction}
    dismissButtonAriaLabel="Close"
  >
    {props.text}
  </MessageBar>
);


///////////// APP
export default class App extends React.Component<AppProps, AppState> {

  constructor(props, context) {
    super(props, context);
    this.state = {
      folders: null,
      listProjects: comboBoxBasicOptions,
      listFolders: null,
      listSubFolders: null,
      projectId: 0,
      categoryId: 0,
      subCategoryId: 0,
      btnDisabled: true,
      modalHide: true,
      mbShow: false,
      mbType: MessageBarType.success,
      mbText: '',
      showSuccess: false,
    };
    sp.setup({
      sp: {
        fetchClientFactory: MsalClientSetup({
          auth: {
            authority: "https://login.microsoftonline.com/raonasn.onmicrosoft.com",
            clientId: "305f533e-8fcf-4461-94e9-bc088a69c7d7",
            redirectUri: "https://raonasn.sharepoint.com/psi/SitePages/Home.aspx",
          },
        }, ["https://raonasn.sharepoint.com/psi/.default"]),
      },
    });
  }

  async componentDidMount() {
    // Cargar combo proyectos
    try {
      const web: IWeb = await sp.web();
      const res = await web.lists.getByTitle("Folders").get();
      console.log('Response: ', res);
    } catch (error) {
      console.log('ERROR: ', error);
    }
  }

  click = async () => {
    this.setState({modalHide: false})
  };

  cbOnItemClick = (event: React.FormEvent<IComboBox>, option: IComboBoxOption, cbName: string) => {
    console.log('Current target', event)
    console.log('Option key: ', option.key)
    switch (cbName) {
      case 'projects':
        this.setState({ listFolders: comboBoxBasicOptions });
        this.setState({ listSubFolders: null });
        this.setState({ btnDisabled: true })
        break;

      case 'folder':
        this.setState({ listSubFolders: comboBoxBasicOptions });
        this.setState({ btnDisabled: true })
        break;

      case 'subfolder':
        this.setState({ btnDisabled: false })
        break;

      default:
        break;
    }
  }

  modalAccept = ():any => {
    this.setState({
      modalHide: true, 
      mbShow: true, 
      mbType: MessageBarType.success,
      mbText: 'El archivo ha sido guardado con éxito',
      showSuccess: true
    })
    formContainerStyles = { display: 'none'}
  }

  modalCancel = () => {
    this.setState({
      listProjects: comboBoxBasicOptions, 
      listFolders: null, 
      listSubFolders : null, 
      btnDisabled: true, 
      modalHide: true, 
      mbShow: true, 
      mbType: MessageBarType.warning,
      mbText: 'La operación ha sido cancelada'
    })
  }

  modalClose = () => {
    this.setState({modalHide: true})
  }

  messageBarClose = () => {
    this.setState({mbShow: false})
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        {/* <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" /> */}
        {/* <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p> */}
        { this.state.mbShow && <AddInMessageBar closeFunction={this.messageBarClose} choice={this.state.mbType} text={this.state.mbText} />}
        <div style={formContainerStyles}>
          <ComboBox
            defaultSelectedKey="C"
            label="Proyectos"
            allowFreeform
            autoComplete="on"
            style={cbStyle}
            options={this.state.listProjects}
            onItemClick={(event, option) => this.cbOnItemClick(event, option, 'projects')}
          />
          <ComboBox
            defaultSelectedKey="C"
            label="Carpeta"
            autoComplete="on"
            style={cbStyle}
            disabled={this.state.listFolders ? false : true}
            options={this.state.listFolders}
            onItemClick={(event, option) => this.cbOnItemClick(event, option, 'folder')}
          />
          <ComboBox
            defaultSelectedKey="C"
            label="Sub carpeta"
            autoComplete="on"
            style={cbStyle}
            disabled={this.state.listSubFolders ? false : true}
            options={this.state.listSubFolders}
            onItemClick={(event, option) => this.cbOnItemClick(event, option, 'subfolder')}
          />
          <PrimaryButton
            text="Guardar archivos"
            style={btnStyle}
            onClick={this.click}
            disabled={this.state.btnDisabled} 
          />
          <DialogBasicExample acceptFunction={this.modalAccept} cancelFunction={this.modalCancel} closeFunction={this.modalClose} hide={this.state.modalHide}/>
        </div>
        { this.state.showSuccess && <div style={successMessageStyles}><h1 className="success-header">Ya puedes cerrar el complemento</h1></div>}
      </div>
    );
  }
}
