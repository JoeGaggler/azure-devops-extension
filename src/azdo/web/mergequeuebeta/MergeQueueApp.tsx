import { Header } from "azure-devops-ui/Components/Header/Header";
import { TitleSize } from "azure-devops-ui/Components/Header/Header.Props";
import { IHeaderCommandBarItem } from "azure-devops-ui/Components/HeaderCommandBar/HeaderCommandBar.Props";
import { Page } from "azure-devops-ui/Components/Page/Page";
import React from "react";
import { getAzdoInfo, getGitClient, TenantInfo } from "./azuredevops";
import { GitPullRequestSearchCriteria, PullRequestStatus, PullRequestTimeRangeType } from "azure-devops-extension-api/Git/Git";
//import { GitPullRequestSearchCriteria } from "azure-devops-extension-api/Git/Git";

export interface MergeQueueAppSingleton {
    bearerToken: string;
    appToken: string;
}

interface ReducerState {
}

interface ReducerAction {

}

function reducer(state: ReducerState, action: ReducerAction): ReducerState {
    let next = { ...state };

    if (action) {
        // TODO
    }

    return next;
}

export function MergeQueueApp(p: { singleton: MergeQueueAppSingleton }) {
    let _tenantInfo = React.useRef<TenantInfo>();
    let singleton = React.useRef(p.singleton);
    let gitClient = React.useRef(getGitClient());

    const [_state, _dispatch] = React.useReducer<(state: ReducerState, action: ReducerAction) => ReducerState>(reducer, {
        // TODO: initial reducer state
    })

    // initialize the app
    React.useEffect(() => { init() }, []);
    async function init() {
        try {
            console.log("MQ: init");
            let nextTenantInfo = await getAzdoInfo();
            if (!nextTenantInfo) {
                // TODO: lock the app
                console.error("Failed to get Azure DevOps info");
                return;
            }
            _tenantInfo.current = nextTenantInfo;
        } catch (error) {
            console.error("MQ: init -> error occurred", error);
        }
    }

    // ticktock
    React.useEffect(() => {
        let id = setInterval(() => { ticktock(); }, 5000);
        return () => clearInterval(id);
    }, []);
    async function ticktock() {
        try {
            console.log("MQ: ticktock");

            if (!singleton.current) {
                console.error("MQ: ticktock -> no singleton available");
                return;
            }

            let tenantInfo = _tenantInfo.current;
            if (!tenantInfo) {
                console.error("MQ: ticktock -> no tenant info available");
                return;
            }

            let git = gitClient.current;
            let criteria: GitPullRequestSearchCriteria = {
                creatorId: undefined!,
                includeLinks: false,
                maxTime: undefined!,
                minTime: undefined!,
                queryTimeRangeType: PullRequestTimeRangeType.Created,
                repositoryId: undefined!,
                reviewerId: undefined!,
                sourceRefName: undefined!,
                sourceRepositoryId: undefined!,
                status: PullRequestStatus.Active,
                targetRefName: undefined!,
                title: undefined!
            };
            let prs = await git.getPullRequestsByProject(tenantInfo.project, criteria, undefined, undefined, undefined);
            console.log("MQ: pull requests", prs);
        } catch (error) {
            console.error("MQ: ticktock -> error occurred", error);
        }
    }

    function renderPageCommandBarItems(): IHeaderCommandBarItem[] {
        // TODO: return page command bar items
        return [];
    }

    return (
        <Page>
            <Header
                title="Merge Queue Beta"
                titleSize={TitleSize.Large}
                commandBarItems={renderPageCommandBarItems()}
            />

            <div className="text-neutral-30 flex-row padding-4">
                <div className="flex-grow"></div>
                <div>__MERGEQUEUEVERSION__</div>
            </div>
        </Page>
    );
}