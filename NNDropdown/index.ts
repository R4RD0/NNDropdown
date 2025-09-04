import * as ReactDOM from 'react-dom';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as operations from './operations';

export class NNDropdown implements ComponentFramework.StandardControl<IInputs, IOutputs> {

  private _context: ComponentFramework.Context<IInputs>;
  private _container: HTMLDivElement;
  private _notifyOutputChanged: () => void;

  // mirrored labels → written to bound text column (boundField)
  private _mirrorValue: string = "";

  constructor() { }

  public async init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ) {
    this._context = context;
    this._container = container;
    this._notifyOutputChanged = notifyOutputChanged;

    operations._execute(this._context, this._container, (labels: string) => {
      this._mirrorValue = labels;
      this._notifyOutputChanged();
    });
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;

    operations._execute(this._context, this._container, (labels: string) => {
      this._mirrorValue = labels;
      this._notifyOutputChanged();
    });
  }

  public getOutputs(): IOutputs {
    // IMPORTANT: must match your manifest property name. You showed "boundField".
    return {
      boundField: this._mirrorValue
    };
  }

  public destroy(): void {
    ReactDOM.unmountComponentAtNode(this._container);
  }
}