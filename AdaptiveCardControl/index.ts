import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as AdaptiveCards from "adaptivecards";
import * as ACData from "adaptivecards-templating";


export class AdaptiveCardControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _notifyOutputChanged: () => void;
	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _card: HTMLElement;
	private _responseData: string;

	constructor() { }

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this._notifyOutputChanged = notifyOutputChanged;
		this._context = context;
		this._container = container;
	}

	public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {

		this._context = context;

		if (this._card) {
			this._container.removeChild(this._card);
		}

		if (this._context.updatedProperties.includes("responseJson")) {
			this.clearResponse();
			this.renderResponse('Thank you for submitting!');			
			this.renderResponse('<br/>');
			await this.handleValue(this._context.parameters.responseJson.raw);
			//this.renderResponse(this._context.parameters.responseJson.raw);
		}
		else if (this._context.parameters.responseJson.raw) {
			this.clearResponse();
			this.renderResponse('Thank you for submitting!');
			this.renderResponse('<br/>');
			//this.renderResponse(this._context.parameters.responseJson.raw);
			//await this.handleValue(this._context.parameters.responseJson.raw);
		}
		else {
			const adaptiveCardTemplateJson = await this.getAdaptiveCardJson(
				this._context.parameters.templateRetrieveEntityType.raw,
				this._context.parameters.templateRetrieveOptions.raw,
				this._context.parameters.templateRetrieveAttributeName.raw,
				this._context.parameters.templateRetrieveFetchXml.raw);

			const adaptiveCardDataJson = await this.getAdaptiveCardJson(
				this._context.parameters.dataRetrieveEntityType.raw,
				this._context.parameters.dataRetrieveOptions.raw,
				this._context.parameters.dataRetrieveAttributeName.raw,
				this._context.parameters.dataRetrieveFetchXml.raw);

			this.renderRequestCard(adaptiveCardTemplateJson, adaptiveCardDataJson);
		}
	}
	
	private isJSON(input: string) { 
		try { 
			return (JSON.parse(input) && !!input); 
		} catch (e) { 
			return false; 
		}
	} 

	private async handleValue(input: string|null)
	{
		if(this.isJSON(input)){
			var json = JSON.parse(input);

			for (const [question, answer] of Object.entries(json))
			{
				//this.renderResponse(json[question]);
				//const title = await this.getQuestionTitle(question);	
				const json = await this.getQuestionTitle(question);	
				const title = json.title;
				const type = json.type;
				var answer_text = answer;
				if (type == "Single Choice")
				{
					// get answer text based on value
					answer_text = await this.getChoiceText(String(answer));	
				}
				this.renderResponse('<p><span style="font-weight: bold">' + title + '</span></p><p>' + answer_text + '</p>');

			}				
		}
	}

	private async getChoiceText(choiceId: string) {
		const fetchXml = '?$select=ki_name&$filter=(ki_questionchoiceid eq ' + choiceId + ')';
		const response = await this._context.webAPI.retrieveMultipleRecords("ki_questionchoice", fetchXml);
		
		if (response && response.entities.length == 1) {
			const text = response.entities[0]["ki_name"];
			return text;
		}

		return null;

	}
	private async getQuestionTitle(questionId: string) {
		const response = await this._context.webAPI.retrieveRecord("ki_question", questionId);

		//if (response && response.entities.length == 1) {			
		if (response) {
			//const title = response["ki_name"];
			//const questionType = response["ki_data_type"];
			var json = {
				"title": response["ki_name"],
				"type": response["ki_data_type"]
			};
		
			return json;
		}

		return null;

	}
	private async getAdaptiveCardJson(entityType: string, options: string, attributeName: string, fetchXml?: string): Promise<any> {

		// @ts-ignore
		const entityId = this._context.page.entityId;
		const retrieveOptions = (fetchXml ? "?fetchXml=" + fetchXml : options || "").replace("{entityId}", entityId);
		const response = await this._context.webAPI.retrieveMultipleRecords(entityType, retrieveOptions);
		
		//alert("entityId: " + entityId);	
		//alert("retrieveOptions: " + retrieveOptions);	
		console.log("entityId: " + entityId);		
		console.log("entityType: " + entityType);		
		console.log("attributeName: " + attributeName);		
		console.log("fetchXml: " + fetchXml);


		if (response && response.entities.length == 1) {
			const adaptiveCardJson = response.entities[0][attributeName];
			const json = JSON.parse(adaptiveCardJson);
			return json;
		}

		return null;
	}

	private renderRequestCard(templateJson: any, dataJson: any) {
		if (!templateJson || !dataJson) return;

		const template = new ACData.Template(templateJson);
		const cardPayload = template.expand({
			$root: dataJson
		});

		const adaptiveCard = new AdaptiveCards.AdaptiveCard();
		adaptiveCard.parse(cardPayload);

		adaptiveCard.onExecuteAction = this.submitRequest.bind(this);

		this._card = adaptiveCard.render();
		this._container.appendChild(this._card);
	}

	private renderResponse(message: string) {
		this._card = document.createElement("div");
		this._card.innerHTML = `<div>${message}</div>`;
		this._container.appendChild(this._card);
	}
	
	private clearResponse() {
		this._container.innerHTML = "";
	}

	private submitRequest(action: AdaptiveCards.Action) {
		const submitAction = action as AdaptiveCards.SubmitAction;
		if (submitAction) {
			this._responseData = JSON.stringify(submitAction.data);
			this._notifyOutputChanged();
		}
	}

	public getOutputs(): IOutputs {
		return {
			responseJson: this._responseData
		};
	}

	public destroy(): void {
	}
}