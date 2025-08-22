// import React from "react";
// import * as joe from "../lib.ts";
// import * as SDK from 'azure-devops-extension-sdk';
// import * as Azdo from '../azdo/azdo.ts';
// import * as luxon from 'luxon'
// import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
// import { Button } from "azure-devops-ui/Button";
// import { Card } from "azure-devops-ui/Card";
// import { Dropdown } from "azure-devops-ui/Dropdown";
// import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";
// import { Icon, IconSize } from "azure-devops-ui/Icon";
// import { IListBoxItem } from "azure-devops-ui/ListBox";
// import { ListSelection } from "azure-devops-ui/List";
import { Page } from "azure-devops-ui/Page";
// import { Pill, PillVariant, PillSize } from "azure-devops-ui/Pill";
// import { PillGroup } from "azure-devops-ui/PillGroup";
// import { ScrollableList, IListItemDetails, ListItem } from "azure-devops-ui/List";
// import { Toast } from "azure-devops-ui/Toast";
// import { Toggle } from "azure-devops-ui/Toggle";
// import { VssPersona } from "azure-devops-ui/VssPersona";
// import { type IHostNavigationService } from 'azure-devops-extension-api';

interface AppSingleton {
    // repositoryFilterDropdownMultiSelection: DropdownMultiSelection;
}

interface AppProps {
    bearerToken: string;
    appToken: string;
    singleton: AppSingleton;
}

function App(p: AppProps) {
    console.log("App render", p);
    return (
        <Page>
            <div className="padding-8 margin-8">
                WIP
            </div>
        </Page>
    )
}

export { App };
export type { AppProps, AppSingleton };
