import ReactDOM from 'react-dom'
import * as SDK from 'azure-devops-extension-sdk';
import { NextRunTab } from './NextRunTab.tsx'
import type { NextRunTabSingleton } from './NextRunTab.tsx'
import { SurfaceBackground, SurfaceContext, Spacing } from "azure-devops-ui/Surface";

console.log("pingmint menu is loading");

SDK.init();

// singleton

const appSingleton: NextRunTabSingleton = {
    appToken: "",
    bearerToken: "",
    build: undefined,
    definition: undefined,
};

let render = () => {
    console.log("render");
    ReactDOM.render(
        <SurfaceContext.Provider value={{
            background: SurfaceBackground.neutral,
            spacing: Spacing.default
        }}>
            <NextRunTab singleton={appSingleton} />
        </SurfaceContext.Provider>,
        document.getElementById('extension_root_div')
    );
}

let refreshMs = 1000 * 60 * 5; // 5 minutes

let refreshToken = () => {
    console.log("refreshToken");
    SDK.getAccessToken().then((token) => {
        console.log("Refreshed token", token);
        appSingleton.bearerToken = token;
        // TODO: also refresh app token
        // render({ bearerToken: token, appToken: "TODO_REFRESH_APP_TOKEN", singleton: appSingleton });
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
        appSingleton.appToken = a;

        SDK.getAccessToken().then((b) => {
            console.log("BearerToken is ready");
            console.log(b);
            appSingleton.bearerToken = b;

            let conf = SDK.getConfiguration();
            console.log("conf", conf);
            SDK.notifyLoadSucceeded();

            SDK.getService("ms.vss-build-web.build-page-data-service").then((buildPageService: any) => {
                const getBuildPageData = buildPageService.getBuildPageData;
                if (getBuildPageData) {
                    getBuildPageData().then((buildPageData: any) => {
                        console.log("main -> Build page data:", buildPageData);
                        if (buildPageData.build && buildPageData.definition) {
                            console.log("Current build is", buildPageData.build);
                            console.log("Current definition is", buildPageData.definition);
                            appSingleton.build = buildPageData.build;
                            appSingleton.definition = buildPageData.definition;
                            render();
                        }
                    });
                }
            });

            setTimeout(refreshToken, refreshMs);
        });
    });
});
