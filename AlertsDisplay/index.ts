import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class AlertsDisplay implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	
	private controlContainer: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	
	/**
	 * Empty constructor.
	 */
	constructor() {
	
	}
	
	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
	
		this._context = context;
		this.controlContainer = document.createElement("div");
		let ul = document.createElement('ul');
	
		container.appendChild(this.controlContainer);
		this.controlContainer.appendChild(ul);
	
		let obj = JSON.parse(this._context.parameters.JSONAlert.raw?.toString()!);

		const myJSON = JSON.stringify(obj);
		let parsedJSON = JSON.parse(myJSON);
		let Object = this;

		this.GetFields(parsedJSON, Object, ul);
	}
	
	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
	}
	
	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}
	
	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
	
	//Reads the JSON text supplied to the PCF componenet and tries to fetch the field and display information
	public async GetFields(parsedJSON: any, Object: this, ul: HTMLUListElement) {
		for (const item of parsedJSON) {
			if (item.hardcodedfunctionname == null) {

				var RelatedLookupObject = await this.getRelatedLookupField(item.lookupfield, item.relatedlookupfield, Object);
				if (RelatedLookupObject != "error") {
					var MetaDataObject = await this.getFieldMetaData(RelatedLookupObject.lookupLogicalName, item.field, Object);
					if (MetaDataObject != "error") {
						var GetDataField = await this.getDataFields(RelatedLookupObject.lookupid, RelatedLookupObject.lookupLogicalName, item.field, item.alert, item.showalertbooleanwhen, item.showalertoptionsetwhen, MetaDataObject.attributetype, Object);
						if (GetDataField != "error") {
							this.displayFields(ul, GetDataField, item.backgroundcolour, item.textcolour);
						}
					}
				}
			}
		}
	}
	
	//Fetches and returns the Lookup properties if the JSON object supplied is looking two entities deep
	public getRelatedLookupField = async (lookupfield: any, relatedlookupfield: any, Object: this): Promise<any> => {
		return new Promise(async function (resolve, reject) {
			const recordId = (<any>Object._context.mode).contextInfo.entityId;
			const lookupLogicalName = (<any>Object._context.mode).contextInfo.entityTypeName;

			if (lookupfield == null)
			{
				resolve({ lookupid: recordId, lookupLogicalName: lookupLogicalName });
			}
			else
			{
			const fieldFormatted = "_" + lookupfield + "_value";
	
			const data = await Object._context.webAPI.retrieveRecord(
				lookupLogicalName,
				recordId,
				"?$select=" + fieldFormatted
			)

			if (data) {
				let recordId = data[fieldFormatted];
				let lookupLogicalName = data[fieldFormatted + "@Microsoft.Dynamics.CRM.lookuplogicalname"];
	
				if (recordId == null) {
					resolve("error");
				}
	
				if (relatedlookupfield != null) {
					const fieldFormatted2 = "_" + relatedlookupfield + "_value";

					const data2 = await Object._context.webAPI.retrieveRecord(
						lookupLogicalName,
						recordId,
						"?$select=" + fieldFormatted2
					)
					if (data2) {
						let recordId = data2[fieldFormatted2];
						let lookupLogicalName = data2[fieldFormatted2 + "@Microsoft.Dynamics.CRM.lookuplogicalname"];

						if (recordId == null) {
							resolve("error");
						}
						else {
							resolve({ lookupid: recordId, lookupLogicalName: lookupLogicalName });
						}
					}
					else {
						resolve("error");
					}
				}
				else {
					resolve({ lookupid: recordId, lookupLogicalName: lookupLogicalName });
				}
			}
			else {
				resolve("error");
			}
		}
		})
	}
	
	//Fetches and returns the attribute type of the field
	public getFieldMetaData = async (LookupLogicalName: any, field: any, Object: this): Promise<any> => {
		return new Promise(function (resolve, reject) {
			var req = new XMLHttpRequest();
			req.open("GET", (<any>Object._context).page.getClientUrl() + "/api/data/v9.1/EntityDefinitions(LogicalName='" + LookupLogicalName + "')/Attributes(LogicalName='" + field + "')", true)
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.setRequestHeader("OData-MaxVersion", "4.0");
			req.setRequestHeader("OData-Version", "4.0");
			req.onreadystatechange = function () {
				if (this.readyState == 4) {
					req.onreadystatechange = null;
						if (this.status == 200) {
						var data = JSON.parse(this.response);
						var attributetype = data.AttributeType;
						resolve({ attributetype: attributetype })
					} else {
							var error = JSON.parse(this.response).error;
						resolve("error");
					}
				}
			};
			req.send();
		});
	}
	
	//Fetches and returns the field formatted value
	public getDataFields = async (lookupid: any, LookupLogicalName: any, field: any, alert: any, showalertbooleanwhen: any, showalertoptionsetwhen: any, attributetype: any, Object: this): Promise<any> => {
		return new Promise(async function (resolve, reject) {
			let fieldvalue = attributetype == "Lookup" ? "_" + field + "_value" : field
	
			const data = await Object._context.webAPI.retrieveRecord(
				LookupLogicalName,
				lookupid,
				"?$select=" + fieldvalue
			)
			//https://docs.microsoft.com/en-us/dynamics365/customer-engagement/web-api/attributetypecode?view=dynamics-ce-odata-9
			if (data) {
				switch (attributetype) {
					case "Boolean": {
						let fieldvalue = data[field];
						let newStr = alert.replace('%field%', fieldvalue)
	
						if (showalertbooleanwhen == fieldvalue) {
							resolve(newStr);
						}
						else {
							resolve("error");
						}
						break;
					}
					case "String": {
						let fieldvalue = data[field];
						let newStr = alert.replace('%field%', fieldvalue);
						if (fieldvalue == null)
						{
							resolve("error");
						}
						else{
							resolve(newStr);
						}
						break;
					}
					case "Picklist": {
						let fieldvalue = data[field + "@OData.Community.Display.V1.FormattedValue"];
						let newStr = alert.replace('%field%', fieldvalue);
						let OptionSetValue = data[field];
						if (fieldvalue == "Yes" || fieldvalue == "No") {
							if (fieldvalue == "No" && showalertbooleanwhen == false) {
								resolve(newStr);
							}
							else if (fieldvalue == "Yes" && showalertbooleanwhen == true) {
								resolve(newStr);
							}
							else {
								resolve("error");
							}
						}
						else if (showalertoptionsetwhen != null) {
							if (showalertoptionsetwhen.includes(OptionSetValue)) {
								resolve(newStr);
							}
							else {
								resolve("error");
							}
						}
						else if (fieldvalue == null) {
							resolve("error");
						}
						else {
							let fieldvalue = data[field + "@OData.Community.Display.V1.FormattedValue"];
							let newStr = alert.replace('%field%', fieldvalue);
							resolve(newStr);
						}
						break;
					}
					case "Lookup": {
						let fieldvalue = data["_" + field + "_value@OData.Community.Display.V1.FormattedValue"];
						let newStr = alert.replace('%field%', fieldvalue);
						resolve(newStr);
						break;
					}
					default: {
						let fieldvalue = data[field + "@OData.Community.Display.V1.FormattedValue"];
						let newStr = alert.replace('%field%', fieldvalue);
						resolve(newStr);
						break;
					}
				}
			}
			else {
				resolve("error");
			}
		})
	}
	
	//Displays the alert onto a HTML List
	public displayFields(ul: HTMLUListElement, newStr: String, backgroundcolour: any, textcolour: any): void {
		let li = document.createElement('li');
		ul.appendChild(li);
		li.innerHTML += newStr;
		li.style.setProperty("background", backgroundcolour);
		li.style.setProperty("color", textcolour);
	}

}
	
	