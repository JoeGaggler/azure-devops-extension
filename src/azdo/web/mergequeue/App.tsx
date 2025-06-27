import React from "react";
// import { CommonServiceIds, type IProjectPageService } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from "azure-devops-extension-api";
// import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
import { getAzdo } from '../azdo/azdo.ts';
import { PullRequestList } from './PullRequestList.tsx';
// import { Button } from "azure-devops-ui/Button";
// import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
// import { TrainCard } from "./TrainCard.tsx"
// import { Dropdown } from "azure-devops-ui/Dropdown";
import { Icon } from "azure-devops-ui/Icon";
import { type IExtensionDataService } from 'azure-devops-extension-api';
// import { ExtensionManagementRestClient } from "azure-devops-extension-api/ExtensionManagement";
import { Toggle } from "azure-devops-ui/Toggle";

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
    const [allPullRequests, setAllPullRequests] = React.useState<Array<any>>([]);
    const [queuedPullRequests, setQueuedPullRequests] = React.useState<Array<any>>([]);
    const [showingDrafts, setShowingDrafts] = React.useState<boolean>(false);

    let mergeQueueDocumentCollectionId = "mergeQueue";
    let mergeQueueDocumentId = "primaryQueue";
    let userPullRequestFiltersDocumentId = "userPullRequestFilters";

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

        setAllPullRequests(pullRequests.value);

        let queue = [
            // {
            //     isBundle: true,
            //     bundle: [
            //         pullRequests.value[0],
            //         pullRequests.value[1]
            //     ]
            // },
            pullRequests.value[0],
            pullRequests.value[1],
            pullRequests.value[2],
            pullRequests.value[3],
            pullRequests.value[4]
        ]
        setQueuedPullRequests(queue);

        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

        let mergeQueueDoc = {
            id: mergeQueueDocumentId,
            prs: []
        }
        console.log("mergeQueueDoc 1: ", mergeQueueDoc);
        try {
            mergeQueueDoc = await dataManager.getDocument(mergeQueueDocumentCollectionId, mergeQueueDocumentId);
            console.log("mergeQueueDoc 2: ", mergeQueueDoc);
        } catch {
            // mergeQueueDoc = {
            //     id: mergeQueueDocumentId,
            //     prs: []
            // }
            mergeQueueDoc = await dataManager.createDocument(mergeQueueDocumentCollectionId, mergeQueueDoc);
            console.log("mergeQueueDoc 3: ", mergeQueueDoc);
        }

        try {
            mergeQueueDoc = await dataManager.updateDocument(mergeQueueDocumentCollectionId, mergeQueueDoc);
            console.log("mergeQueueDoc 4: ", mergeQueueDoc);
        }
        catch {
        }

        let userFiltersDoc = {
            id: userPullRequestFiltersDocumentId,
            drafts: false,
        }
        console.log("userFiltersDoc 1: ", mergeQueueDoc);
        try {
            userFiltersDoc = await dataManager.getDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, { scopeType: "User" });
            console.log("userFiltersDoc 2: ", userFiltersDoc);
        } catch {
            userFiltersDoc = await dataManager.createDocument(mergeQueueDocumentCollectionId, userFiltersDoc, { scopeType: "User" });
            console.log("userFiltersDoc 3: ", userFiltersDoc);
        }

        setShowingDrafts(userFiltersDoc.drafts);

        try {
            userFiltersDoc = await dataManager.updateDocument(mergeQueueDocumentCollectionId, userFiltersDoc, { scopeType: "User" });
            console.log("userFiltersDoc 4: ", userFiltersDoc);
        }
        catch {
        }
    }

    async function persistShowingDrafts(value: boolean) {
        setShowingDrafts(value);

        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

        let userFiltersDoc = {
            id: userPullRequestFiltersDocumentId,
            drafts: value,
        }
        console.log("userFiltersDoc 5: ", userFiltersDoc);
        try {
            userFiltersDoc = await dataManager.getDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, { scopeType: "User" });
            console.log("userFiltersDoc 6: ", userFiltersDoc);
        } catch {
            userFiltersDoc = await dataManager.createDocument(mergeQueueDocumentCollectionId, userFiltersDoc, { scopeType: "User" });
            console.log("userFiltersDoc 7: ", userFiltersDoc);
        }

        userFiltersDoc.drafts = value;

        try {
            userFiltersDoc = await dataManager.updateDocument(mergeQueueDocumentCollectionId, userFiltersDoc, { scopeType: "User" });
            console.log("userFiltersDoc 8: ", userFiltersDoc);
        }
        catch {
        }
    }

    return (
        <>
            <div className="padding-8 margin-8">

                <h2>Merge Queue</h2>
                <Card className="padding-8">
                    {/* <ButtonGroup className="flex-wrap">
                        <Button
                            text="New Train"
                            onClick={() => alert("TODO: Create a new release train")}
                        />
                    </ButtonGroup> */}
                    <PullRequestList
                        pullRequests={queuedPullRequests}
                        organization={org}
                        project={proj}
                        filters={{}}
                    />
                </Card>

                <br />

                <h2>All Pull Requests</h2>
                <Card className="padding-8">
                    <div className="flex-column">
                        <div className="padding-8 flex-row">
                            <Toggle
                                offText={"Hiding drafts"}
                                onText={"Showing drafts"}
                                checked={showingDrafts}
                                onChange={(_event, value) => { persistShowingDrafts(value); }}
                            />
                        </div>
                        {/* <div className="flex-row">
                        <Dropdown
                            items={[
                                { id: "include", text: "Include drafts" },
                                { id: "exclude", text: "Exclude drafts" },
                                { id: "only", text: "Only drafts" }
                            ]}
                            onSelect={() => { }}
                            placeholder="Select drafts"
                        />
                        <Dropdown
                            items={[
                                { id: "include", text: "Include drafts" },
                                { id: "exclude", text: "Exclude drafts" },
                                { id: "only", text: "Only drafts" }
                            ]}
                            onSelect={() => { }}
                            placeholder="Demo dropdown"
                        />
                    </div> */}
                        <PullRequestList
                            pullRequests={allPullRequests}
                            organization={org}
                            project={proj}
                            filters={
                                (() => {
                                    return {
                                        drafts: showingDrafts
                                    }
                                })()
                            }
                        />
                    </div>
                </Card>

                <Icon iconName="Video" />

            </div>
        </>
    )
}

export { App };
export type { AppProps };
