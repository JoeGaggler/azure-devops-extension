import ReactDOM from 'react-dom'
import * as SDK from 'azure-devops-extension-sdk';
import { App } from './App.tsx'
import type { AppProps } from './App.tsx'

console.log("pingmint menu is loading");

SDK.init();

let render = (p: AppProps) => {
    console.log("render");
    ReactDOM.render(
        <App appToken={p.appToken} bearerToken={p.bearerToken} />,
        document.getElementById('extension_root_div')
    );
}

let refreshMs = 1000 * 60 * 5; // 5 minutes

let refreshToken = () => {
    console.log("refreshToken");
    SDK.getAccessToken().then((token) => {
        console.log("Refreshed token", token);
        render({ bearerToken: token, appToken: "TODO_REFRESH_APP_TOKEN" });
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
            render({ bearerToken: b, appToken: a });
            SDK.notifyLoadSucceeded();

            setTimeout(refreshToken, refreshMs);
        });
    });
});
