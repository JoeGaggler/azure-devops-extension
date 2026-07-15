import { TitleSize } from "azure-devops-ui/Header";
import { Panel } from "azure-devops-ui/Panel";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
import { useState } from "react";

export function AddMergeQueuePanel(p: AddMergeQueuePanelProps) {
    const [id, setId] = useState("")
    const [name, setName] = useState("")

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
            <div className="flex-column">
                <TextField
                    label={"Name"}
                    value={name}
                    onChange={(e, nextValue) => e && setName(nextValue)}
                    width={TextFieldWidth.standard}
                />

                <TextField
                    label={"Id"}
                    value={id}
                    onChange={(e, nextValue) => e && setId(nextValue)}
                    width={TextFieldWidth.standard}
                />
            </div>
    </Panel>
}

export interface AddMergeQueuePanelProps {
    onCommit: (data: AddMergeQueuePanelValues) => void
    onCancel: () => void
}

export interface AddMergeQueuePanelValues {

}
