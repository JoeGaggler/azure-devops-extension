import React from "react";
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";
// import * as SDK from 'azure-devops-extension-sdk';

export interface NextRunTabSingleton {
    bearerToken: string;
    appToken: string;
    
    build: any;
    definition: any;
}

export interface NextRunTabProps {
    singleton: NextRunTabSingleton;
}

export function NextRunTab(p: NextRunTabProps) {
    console.log("NextRunTab render", p);

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        console.log("NextRunTab -> init");
    }

    React.useEffect(() => {
        let id = setInterval(() => {tick();}, 5000);
        return () => clearInterval(id);
    }, []);
    async function tick() {
        console.log("NextRunTab -> tick");
    }

    return <>
        <MessageCard severity={MessageCardSeverity.Info}>
            Work in progress. Please check back later.
        </MessageCard>

        <p>
            __NEXTRUNVERSION__
        </p>
    </>
}