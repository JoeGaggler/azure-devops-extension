// import { CommonServiceIds, type IProjectPageService } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from "azure-devops-extension-api";
import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";

interface AppProps {
    bearerToken: string | null;
    appToken: string | null;
}

function App(p: AppProps) {
    if (p) {
        console.log("App props:", p);
    }

    async function doit() {
        // const navService = await SDK.getService<IHostNavigationService>("ms.vss-features.host-navigation-service");
        // const proj = navService.getPageRoute()

        const projectInfoService = await SDK.getService<IProjectPageService>(
            "ms.vss-tfs-web.tfs-page-data-service" // TODO: CommonServiceIds.ProjectPageService
        );
        const proj = await projectInfoService.getProject();
        // // let proj = "11"
        console.log("Project:", proj);
    }

    return (
        <>
            <Card className="padding-8">
                <div className="flex-row rhythm-horizontal-8 flex-center">
                    <div>Coming soon.</div>
                </div>
                <Button
                    text="Debug"
                    primary={true}
                    onClick={() => { doit(); }}
                />
            </Card>
        </>
    )
}

export { App };
export type { AppProps };
