declare interface IMenuCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'MenuCommandSetStrings' {
  const strings: IMenuCommandSetStrings;
  export = strings;
}
