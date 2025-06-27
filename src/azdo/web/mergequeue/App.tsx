import React from "react";
import * as SDK from 'azure-devops-extension-sdk';
import { Card } from "azure-devops-ui/Card";
import * as Azdo from '../azdo/azdo.ts';
import { PullRequestList } from './PullRequestList.tsx';
import { Icon } from "azure-devops-ui/Icon";
import { type IExtensionDataService } from 'azure-devops-extension-api';
import { Toggle } from "azure-devops-ui/Toggle";

interface AppProps {
    bearerToken: string | null;
    appToken: string | null;
}

interface PullRequestFilters {
    drafts: boolean;
    allBranches: boolean;
}

interface ExtensionDocument {
    id: string;
    __etag?: number;
}

interface MergeQueueList extends ExtensionDocument {
    queues: Array<MergeQueue>;
}

interface MergeQueue {
    pullRequests: Array<any>;
}

function App(p: AppProps) {
    if (p) {
        console.log("App props:", p);
    }

    // Extension Document IDs
    let mergeQueueDocumentCollectionId = "mergeQueue";
    let primaryQueueDocumentId = "primaryQueue";
    let mergeQueueListDocumentId = "mergeQueueList";
    let repoCacheDocumentId = "repoCache";
    let userPullRequestFiltersDocumentId = "userPullRequestFilters";

    const [tenantInfo, setTenantInfo] = React.useState<Azdo.TenantInfo>({});
    const [allPullRequests, setAllPullRequests] = React.useState<Array<any>>([]);
    const [filters, setFilters] = React.useState<PullRequestFilters>({ drafts: false, allBranches: false });
    const [repoMap, setRepoMap] = React.useState<any>({});
    const [mergeQueueList, setMergeQueueList] = React.useState<MergeQueueList>({ queues: [], id: mergeQueueListDocumentId });


    React.useEffect(() => { init() }, []); // run once
    async function init() {
        let bearer = await SDK.getAccessToken()

        let info = await Azdo.getAzdoInfo();
        setTenantInfo(info)

        let pullRequests = await Azdo.getAzdo(`https://dev.azure.com/${info.organization}/${info.project}/_apis/git/pullrequests?api-version=7.2-preview.2`, bearer as string);
        console.log("Pull Requests value:", pullRequests.value);

        refreshRepos(pullRequests.value); // not awaited

        setAllPullRequests(pullRequests.value);

        // setup merge queue list
        let newMergeQueueList: MergeQueueList = {
            id: mergeQueueListDocumentId,
            queues: []
        }
        newMergeQueueList = await Azdo.getOrCreateSharedDocument(mergeQueueDocumentCollectionId, mergeQueueListDocumentId, newMergeQueueList)
        await Azdo.trySaveSharedDocument(mergeQueueDocumentCollectionId, mergeQueueListDocumentId, newMergeQueueList); // TODO: REMOVE THIS
        newMergeQueueList = {
            id: mergeQueueListDocumentId,
            queues: [
                {
                    pullRequests: [
                        // HACK: for demo purposes, we will just take the first 5 pull requests
                        pullRequests.value[0],
                        pullRequests.value[1],
                        pullRequests.value[2],
                        pullRequests.value[3],
                        pullRequests.value[4]
                    ]
                }
            ]
        }
        setMergeQueueList(newMergeQueueList);

        // setup user filters
        let userFiltersDoc = {
            drafts: false,
            allBranches: false
        }
        userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc)
        setFilters({ ...userFiltersDoc });
    }

    // TODO: FIX THIS
    function getPrimaryPullRequests(): Array<any> {
        if (mergeQueueList && mergeQueueList.queues && mergeQueueList.queues.length > 0) {
            return mergeQueueList.queues[0].pullRequests;
        }
        return [];
    }

    async function refreshRepos(value: any) {
        let map = repoMap;

        let sharedMap = await Azdo.getOrCreateSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, {});
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
                let repo = await Azdo.getAzdo(pullRequest.repository.url, p.bearerToken as string);
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
        Azdo.trySaveSharedDocument(mergeQueueDocumentCollectionId, repoCacheDocumentId, sharedMap);

        console.log("Repo map:", map);
        console.log("Shared map:", sharedMap);
    }

    async function persistFilters(value: PullRequestFilters) {
        setFilters(value);

        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>("ms.vss-features.extension-data-service");
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

        let userFiltersDoc = { ...value };
        userFiltersDoc = await Azdo.getOrCreateUserDocument(mergeQueueDocumentCollectionId, userPullRequestFiltersDocumentId, userFiltersDoc);

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
                        pullRequests={getPrimaryPullRequests()}
                        organization={tenantInfo.organization}
                        project={tenantInfo.project}
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
                            organization={tenantInfo.organization}
                            project={tenantInfo.project}
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
