import React from "react";
// import { CommonServiceIds, type IProjectPageService } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from "azure-devops-extension-api";
// import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
import { getAzdo } from '../azdo/azdo.ts';
import { PullRequestList } from './PullRequestList.tsx';
import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import { TrainCard } from "./TrainCard.tsx"


interface AppProps {
    bearerToken: string | null;
    appToken: string | null;
}

function App(p: AppProps) {
    if (p) {
        console.log("App props:", p);
    }

    const [org, setOrg] = React.useState<string | undefined>();
    const [proj, setProj] = React.useState<string | undefined>();
    const [allPullRequests, setAppPullRequests] = React.useState<Array<any>>([]);

    // run once
    React.useEffect(() => { go() }, []);
    async function go() {
        let bearer = await SDK.getAccessToken()

        let host = SDK.getHost()
        console.log("Host:", host);
        setOrg(host.name);

        const projectInfoService = await SDK.getService<IProjectPageService>(
            "ms.vss-tfs-web.tfs-page-data-service" // TODO: CommonServiceIds.ProjectPageService
        );
        const proj = await projectInfoService.getProject();
        console.log("Project:", proj);
        if (proj) { setProj(proj.name); }

        let pullRequests = await getAzdo(`https://dev.azure.com/${host.name}/${proj?.name}/_apis/git/pullrequests?api-version=7.2-preview.2`, bearer as string);
        console.log("Pull Requests value:", pullRequests.value);

        setAppPullRequests(pullRequests.value);
    }

    return (
        <>
            <div className="padding-8 margin-8">

                <h2>Merge Queue</h2>
                <Card className="padding-8">
                    <ButtonGroup className="flex-wrap">
                        <Button
                            text="New Train"
                            onClick={() => alert("TODO: Create a new release train")}
                        />
                    </ButtonGroup>
                </Card>

                <TrainCard
                    name="Demo Train #1"
                />

                <TrainCard
                    name="Demo Train #2"
                />

                <br />

                <h2>All Pull Requests</h2>
                <Card className="padding-8">
                    <PullRequestList
                        pullRequests={allPullRequests}
                        organization={org}
                        project={proj}
                    />
                </Card>

            </div>
        </>
    )
}

export { App };
export type { AppProps };
