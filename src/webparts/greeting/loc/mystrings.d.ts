declare interface IGreetingWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    DescriptionFieldLabel: string;
}

declare module 'GreetingWebPartStrings' {
    const strings: IGreetingWebPartStrings;
    export = strings;
}
