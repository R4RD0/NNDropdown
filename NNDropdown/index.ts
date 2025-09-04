import * as ReactDOM from 'react-dom';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as operations from './operations';

export class NNDropdown implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _notifyOutputChanged: () => void;

	// NEW: store mirror value to write back into bound text field
	private _mirrorValue: string = "";

	constructor() { }

	/**
	 * Called when the control is initialised.
	 */
	public async init(
		context: ComponentFramework.Context<IInputs>,
		notifyOutputChanged: () => void,
		state: ComponentFramework.Dictionary,
		container: HTMLDivElement
	) {
		console.log('Init Context', context, 'Init State', state, 'Init Container', container);

		this._context = context;
		this._container = container;
		this._notifyOutputChanged = notifyOutputChanged;

		// execute and supply callback to update mirror value
		operations._execute(this._context, this._container, (labels: string) => {
			this._mirrorValue = labels;
			this._notifyOutputChanged();
		});
	}

	/**
	 * Called when context/props change (including data).
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		console.log('updateView Context', context);

		this._context = context;

		operations._execute(this._context, this._container, (labels: string) => {
			this._mirrorValue = labels;
			this._notifyOutputChanged();
		});
	}

	/**
	 * Called by the framework to retrieve outputs.
	 */
	public getOutputs(): IOutputs {
		return {
			// IMPORTANT: match the property name you defined in ControlManifest.Input.xml
			boundField: this._mirrorValue
		};
	}

	/**
	 * Cleanup.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this._container);
	}
}
