/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    responseJson: ComponentFramework.PropertyTypes.StringProperty;
    templateRetrieveEntityType: ComponentFramework.PropertyTypes.StringProperty;
    templateRetrieveOptions: ComponentFramework.PropertyTypes.StringProperty;
    templateRetrieveAttributeName: ComponentFramework.PropertyTypes.StringProperty;
    templateRetrieveFetchXml: ComponentFramework.PropertyTypes.StringProperty;
    dataRetrieveEntityType: ComponentFramework.PropertyTypes.StringProperty;
    dataRetrieveOptions: ComponentFramework.PropertyTypes.StringProperty;
    dataRetrieveAttributeName: ComponentFramework.PropertyTypes.StringProperty;
    dataRetrieveFetchXml: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    responseJson?: string;
}
