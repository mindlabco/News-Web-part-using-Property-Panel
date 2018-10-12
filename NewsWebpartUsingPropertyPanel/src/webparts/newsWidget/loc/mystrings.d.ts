declare interface INewsWidgetWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ShowDate: string = "Show Date";
}

declare module 'NewsWidgetWebPartStrings' {
  const strings: INewsWidgetWebPartStrings;
  export = strings;
}
