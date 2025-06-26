// import { CommonServiceIds, type IProjectPageService } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from "azure-devops-extension-api";
import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
import { getAzdo } from '../azdo/azdo.ts';


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

        let host = SDK.getHost()
        console.log("Host:", host);

        const projectInfoService = await SDK.getService<IProjectPageService>(
            "ms.vss-tfs-web.tfs-page-data-service" // TODO: CommonServiceIds.ProjectPageService
        );
        const proj = await projectInfoService.getProject();
        // // let proj = "11"
        console.log("Project:", proj);

        let pullRequests = await getAzdo(`https://dev.azure.com/${host.name}/${proj?.name}/_apis/git/pullrequests?api-version=7.2-preview.2`, p.bearerToken as string);
        console.log("Pull Requests:", pullRequests);
    }

    return (
        <>
            <Card className="padding-8 margin-8">
                <div className="flex-row rhythm-horizontal-8 flex-center">
                    <div>Coming soon.</div>
                </div>
                <br />
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
