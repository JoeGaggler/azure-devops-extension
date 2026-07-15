import ReactDOM from 'react-dom'
import * as SDK from 'azure-devops-extension-sdk';
import { SurfaceBackground, SurfaceContext, Spacing } from "azure-devops-ui/Surface";
import { MergeQueueApp, MergeQueueAppSingleton } from './MergeQueueApp';

SDK.init();

// singleton

const appSingleton: MergeQueueAppSingleton = {
    appToken: "",
    bearerToken: "",
};

let render = () => {
    ReactDOM.render(
        <SurfaceContext.Provider value={{
            background: SurfaceBackground.neutral,
            spacing: Spacing.default
        }}>
            <MergeQueueApp singleton={appSingleton} />
        </SurfaceContext.Provider>,
        document.getElementById('extension_root_div')
    );
}

let refreshMs = 1000 * 60 * 5; // 5 minutes

let refreshToken = () => {
    SDK.getAccessToken().then((token) => {
        appSingleton.bearerToken = token;
        // TODO: also refresh app token
        // render({ bearerToken: token, appToken: "TODO_REFRESH_APP_TOKEN", singleton: appSingleton });
        setTimeout(refreshToken, refreshMs);
    }).catch((err) => {
        console.error("Error getting access token", err);
    });
}

SDK.ready().then(() => {
    SDK.getAppToken().then((a) => {
        appSingleton.appToken = a;

        SDK.getAccessToken().then((b) => {
            appSingleton.bearerToken = b;

            // let conf = SDK.getConfiguration();
            SDK.notifyLoadSucceeded();

            render();

            setTimeout(refreshToken, refreshMs);
        });
    });
});
