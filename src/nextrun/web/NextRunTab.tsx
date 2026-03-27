
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";

export interface NextRunTabSingleton {
    // repositoryFilterDropdownMultiSelection: DropdownMultiSelection;
}

export interface NextRunTabProps {
    bearerToken: string;
    appToken: string;
    singleton: NextRunTabSingleton;
}

export function NextRunTab(p: NextRunTabProps) {
    console.log("NextRunTab render", p);
    return <>
        <MessageCard severity={MessageCardSeverity.Info}>
            Work in progress. Please check back later.
        </MessageCard>

        <p>
            __NEXTRUNVERSION__
        </p>
    </>
}