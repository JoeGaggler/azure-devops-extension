import ReactDOM from 'react-dom'
import * as SDK from 'azure-devops-extension-sdk';
import { NextRunTab } from './NextRunTab.tsx'
import type { NextRunTabSingleton } from './NextRunTab.tsx'
import { SurfaceBackground, SurfaceContext, Spacing } from "azure-devops-ui/Surface";

SDK.init();

// singleton

const appSingleton: NextRunTabSingleton = {
    appToken: "",
    bearerToken: "",
    build: undefined,
    definition: undefined,
};

let render = () => {
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

            SDK.getService("ms.vss-build-web.build-page-data-service").then((buildPageService: any) => {
                const getBuildPageData = buildPageService.getBuildPageData;
                if (getBuildPageData) {
                    getBuildPageData().then((buildPageData: any) => {
                        if (buildPageData.build && buildPageData.definition) {
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
