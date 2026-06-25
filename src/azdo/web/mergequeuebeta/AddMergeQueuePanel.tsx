import { TitleSize } from "azure-devops-ui/Header";
import { Panel } from "azure-devops-ui/Panel";

export function AddMergeQueuePanel(p: AddMergeQueuePanelProps) {
    function onCancel() { p.onCancel() }
    
    function onCreate() { p.onCommit({}) } // TODO: VALUES

    return <Panel
        onDismiss={() => onCancel()}
        titleProps={
            {
                text: "New Merge Queue",
                iconProps: {
                    iconName: "Add"
                },
                size: TitleSize.Medium,
                className: undefined,
                id: undefined,
            }
        }
        description={undefined}
        footerButtonProps={[
            { text: "Cancel", onClick: () => onCancel(), primary: false, },
            { text: "Create", onClick: () => onCreate(), primary: true, },
        ]}>

    </Panel>
}

export interface AddMergeQueuePanelProps {
    onCommit: (data: AddMergeQueuePanelValues) => void
    onCancel: () => void
}

export interface AddMergeQueuePanelValues {

}
