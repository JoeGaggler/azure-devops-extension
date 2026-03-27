import ReactDOM from 'react-dom'
import * as SDK from 'azure-devops-extension-sdk';
import { NextRunTab } from './NextRunTab.tsx'
import type { NextRunTabProps, NextRunTabSingleton } from './NextRunTab.tsx'
import { SurfaceBackground, SurfaceContext } from "azure-devops-ui/Surface";
// import { DropdownMultiSelection } from "azure-devops-ui/Utilities/DropdownSelection";

console.log("pingmint menu is loading");

SDK.init();

// singleton

const appSingleton: NextRunTabSingleton = {
    // repositoryFilterDropdownMultiSelection: new DropdownMultiSelection()
};

let render = (p: NextRunTabProps) => {
    console.log("render");
    ReactDOM.render(
        <SurfaceContext.Provider value={{ background: SurfaceBackground.neutral }}>
           <NextRunTab appToken={p.appToken} bearerToken={p.bearerToken} singleton={appSingleton} />
        </SurfaceContext.Provider>,
        document.getElementById('extension_root_div')
    );
}

let refreshMs = 1000 * 60 * 5; // 5 minutes

let refreshToken = () => {
    console.log("refreshToken");
    SDK.getAccessToken().then((token) => {
        console.log("Refreshed token", token);
        render({ bearerToken: token, appToken: "TODO_REFRESH_APP_TOKEN", singleton: appSingleton });
        setTimeout(refreshToken, refreshMs);
    }).catch((err) => {
        console.error("Error getting access token", err);
    });
}

SDK.ready().then(() => {
    console.log("SDK is ready");
    SDK.getAppToken().then((a) => {
        console.log("AppToken is ready");
        console.log(a);
        SDK.getAccessToken().then((b) => {
            console.log("BearerToken is ready");
            console.log(b);

            let conf = SDK.getConfiguration();
            console.log("conf", conf);
            render({ bearerToken: b, appToken: a, singleton: appSingleton });
            SDK.notifyLoadSucceeded();

            setTimeout(refreshToken, refreshMs);
        });
    });
});
