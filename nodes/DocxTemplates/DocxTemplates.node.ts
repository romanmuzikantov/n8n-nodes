import { createReport } from 'docx-templates';
import { IBinaryData, IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription, NodeOperationError } from 'n8n-workflow';
import fetch from 'node-fetch';

export class DocxTemplates implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'DocxTemplates',
		name: 'DocxTemplates',
		icon: 'file:httpbin.svg',
		group: ['transform'],
		version: 1,
		description: 'Transforms a docx file\'s content based on a json',
		defaults: {
			name: 'DocxTemplates',
		},
		inputs: ['main'],
		outputs: ['main'],
		properties: [
			{
				displayName: 'File URL',
				description: 'URL Path where we can find and retrieve the document',
				name: 'fileUrl',
				type: 'string',
				required: true,
				default: '',
			},
			{
				displayName: 'JSON Template',
				description: 'JSON representing the template',
				name: 'jsonTemplate',
				type: 'json',
				required: true,
				default: '',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		console.log("test");
		const items = this.getInputData();

		let fileUrl: string;
		let jsonTemplate: string;

		let arrayBuffer: ArrayBuffer;
		let template: Buffer;

		const returnData: INodeExecutionData[] = [];

		// Iterates over all input items and add the key "myString" with the
		// value the parameter "myString" resolves to.
		// (This could be a different value for each item in case it contains an expression)
		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				fileUrl = this.getNodeParameter('fileUrl', itemIndex, '') as string;
				jsonTemplate = this.getNodeParameter('jsonTemplate', itemIndex, '') as string;
				console.log(jsonTemplate);
				console.log("test")
				let test = JSON.parse(jsonTemplate);
				console.log(test);

				arrayBuffer = await fetch(fileUrl).then(res => res.arrayBuffer());
				template = Buffer.from(arrayBuffer);

				const responseBuffer = await createReport({
					template,
					data: {name:"Roman"}
				});

				const binary = { 
					["data"]: {
						data: "",
						fileName: 'fileName',
						mimeType: 'mimeType'
					} as IBinaryData
				};
				binary!['data'] = await this.helpers.prepareBinaryData(Buffer.from(responseBuffer), 'test.docx')

				const json = {};
				const result: INodeExecutionData = {
					json,
					binary
				}

				returnData.push(result);
			} catch (error) {
				if (this.continueOnFail()) {
					items.push({ json: this.getInputData(itemIndex)[0].json, error, pairedItem: itemIndex });
				} else {
					if (error.context) {
						error.context.itemIndex = itemIndex;
						throw error;
					}
					throw new NodeOperationError(this.getNode(), error, {
						itemIndex,
					});
				}
			}
		}

		return this.prepareOutputData(returnData);
	}
}
