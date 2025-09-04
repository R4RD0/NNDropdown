import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as operations from './operations';
import type { Root } from 'react-dom/client';

export class NNDropdown implements ComponentFramework.StandardControl<IInputs, IOutputs> {

  private _context: ComponentFramework.Context<IInputs>;
  private _container: HTMLDivElement;
  private _notifyOutputChanged: () => void;

  // mirrored labels → written to bound text column (must match manifest prop)
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
    return {
      // If you renamed in manifest, change "boundField" to that name (e.g., mirrorText)
      boundField: this._mirrorValue
    };
  }

  public destroy(): void {
    // Unmount React 18 root cleanly
    const root = (operations as any).roots?.get(this._container) as Root | undefined;
    if (root) root.unmount();
  }
}
