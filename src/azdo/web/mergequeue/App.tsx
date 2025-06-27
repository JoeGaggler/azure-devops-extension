import React from "react";
// import { CommonServiceIds, type IProjectPageService } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { IProjectPageService } from "azure-devops-extension-api";
// import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
import { getAzdo, getOrCreateUserDocument, getOrCreateSharedDocument, trySaveSharedDocument } from '../azdo/azdo.ts';
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

interface PullRequestFilters {
    drafts: boolean;
    allBranches: boolean;
}

function App(p: AppProps) {
    if (p) {
        console.log("App props:", p);
    }

    const [org, setOrg] = React.useState<string | undefined>();
    const [proj, setProj] = React.useState<string | undefined>();
    const [allPullRequests, setAllPullRequests] = React.useState<Array<any>>([]);
    const [queuedPullRequests, setQueuedPullRequests] = React.useState<Array<any>>([]);
    const [filters, setFilters] = React.useState<PullRequestFilters>({ drafts: false, allBranches: false });
    const [repoMap, setRepoMap] = React.useState<any>({});

    let mergeQueueDocumentCollectionId = "mergeQueue";
    let mergeQueueDocumentId = "primaryQueue";
    let repoCacheDocumentId = "repoCache";
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

        await refreshRepos(pullRequests.value);

        setAllPullRequests(pullRequests.value);

        let queue = [
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
            prs: []
        }
        mergeQueueDoc = await getOrCreateSharedDocument(mergeQueueDocumentCollectionId, mergeQueueDocumentId, mergeQueueDoc)
        await trySaveSharedDocument(mergeQueueDocumentCollectionId, mergeQueueDocumentId, mergeQueueDoc);

        let userFiltersDoc = {
            drafts: false,
            allBranches: false
        }
        userFiltersDoc = await getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc)

        setFilters({ ...userFiltersDoc });

        try {
            userFiltersDoc = await dataManager.updateDocument(mergeQueueDocumentCollectionId, userFiltersDoc, { scopeType: "User" });
            console.log("userFiltersDoc 4: ", userFiltersDoc);
        }
        catch {
        }
    }

    async function refreshRepos(value: any) {
        let map = repoMap;

        let sharedMap = await getOrCreateSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, {});
        if (sharedMap) {
            console.log("Shared repo map:", sharedMap);
        }

        for (let pullRequest of value) {
            if (pullRequest.repository.name && map[pullRequest.repository.name]) {
                // already cached
                continue
            }
            if (sharedMap && sharedMap[pullRequest.repository.name]) {
                let repo = sharedMap[pullRequest.repository.name];
                if (repo && repo.id && repo.name && repo.defaultBranch) {
                    map[pullRequest.repository.name] = repo;
                    continue
                }
            }
            if (pullRequest.repository && pullRequest.repository.name && pullRequest.repository.url) {
                let repo = await getAzdo(pullRequest.repository.url, p.bearerToken as string);
                console.log("Repo:", pullRequest.repository.name, repo);
                let newRepo = {
                    id: repo.id,
                    name: repo.name,
                    defaultBranch: repo.defaultBranch,
                }
                map[pullRequest.repository.name] = newRepo;
                sharedMap[pullRequest.repository.name] = newRepo;
            } else {
                // console.warn("No repository found for pull request:", pullRequest);
            }
        }

        setRepoMap(map);
        trySaveSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, sharedMap);

        console.log("Repo map:", map);
        console.log("Shared map:", sharedMap);
    }

    async function persistFilters(value: PullRequestFilters) {
        setFilters(value);

        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

        let userFiltersDoc = { ...value };
        userFiltersDoc = await getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);

        userFiltersDoc.drafts = value.drafts;
        userFiltersDoc.allBranches = value.allBranches;

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
                        repos={repoMap}
                    />
                </Card>

                <br />

                <h2>All Pull Requests</h2>
                <Card className="padding-8">
                    <div className="flex-column">
                        <div className="padding-8 flex-row rhythm-horizontal-16">
                            <Toggle
                                offText={"All branches"}
                                onText={"All branches"}
                                checked={filters.allBranches}
                                onChange={(_event, value) => { persistFilters({ ...filters, allBranches: value }) }}
                            />
                            <Toggle
                                offText={"Drafts"}
                                onText={"Drafts"}
                                checked={filters.drafts}
                                onChange={(_event, value) => { persistFilters({ ...filters, drafts: value }) }}
                            />
                        </div>
                        <PullRequestList
                            pullRequests={allPullRequests}
                            organization={org}
                            project={proj}
                            filters={filters}
                            repos={repoMap}
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
