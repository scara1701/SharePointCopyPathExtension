declare interface IMyCustomCommandbarCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'MyCustomCommandbarCommandSetStrings' {
  const strings: IMyCustomCommandbarCommandSetStrings;
  export = strings;
}
