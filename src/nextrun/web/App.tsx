export interface AppSingleton {
    // repositoryFilterDropdownMultiSelection: DropdownMultiSelection;
}

export interface AppProps {
    bearerToken: string;
    appToken: string;
    singleton: AppSingleton;
}

export function App(p: AppProps) {
    console.log("App render", p);
    return <p>hello1</p>
}