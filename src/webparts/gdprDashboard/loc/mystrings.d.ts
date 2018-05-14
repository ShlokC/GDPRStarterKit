declare interface IGdprDashboardWebPartStrings {
  BasicGroupName: string;
  PropertyPaneDescription: string;

  TaskCompletedColumnTitle: string;
  TaskAssigneeColumnTitle: string;
  TaskTitleColumnTitle: string;
  TaskDueDateColumnTitle: string;
}

declare module 'GdprDashboardWebPartStrings' {
  const strings: IGdprDashboardWebPartStrings;
  export = strings;
}
