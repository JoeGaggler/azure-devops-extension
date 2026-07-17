import { TitleSize } from "azure-devops-ui/Header";
import { Panel } from "azure-devops-ui/Panel";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
import { useState } from "react";

export function AddMergeQueuePanel(p: AddMergeQueuePanelProps) {
    const [id, setId] = useState(p.id || "")
    const [name, setName] = useState(p.name || "")
    const [targetRefName, setTargetRefName] = useState(p.targetRefName || "")

    function onCancel() { p.onCancel() }

    function onCreate() {
        p.onCommit({
            id: id,
            name: name,
            targetRefName: targetRefName
        })
    } // TODO: VALUES

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
            { text: "Save", onClick: () => onCreate(), primary: true, },
        ]}>

        <div className="flex-column">
            {
                (p.id === undefined) && (
                    <TextField
                        label={"Id"}
                        value={id}
                        onChange={(e, nextValue) => e && setId(nextValue)}
                        width={TextFieldWidth.standard}
                    />
                )
            }

            <TextField
                label={"Name"}
                value={name}
                onChange={(e, nextValue) => e && setName(nextValue)}
                width={TextFieldWidth.standard}
            />

            <TextField
                label={"Branch"}
                value={targetRefName}
                onChange={(e, nextValue) => e && setTargetRefName(nextValue)}
                width={TextFieldWidth.standard}
            />

        </div>
    </Panel>
}

export interface AddMergeQueuePanelProps {
    id?: string
    name?: string
    targetRefName?: string
    onCommit: (data: AddMergeQueuePanelValues) => void
    onCancel: () => void
}

export interface AddMergeQueuePanelValues {
    id: string
    name: string
    targetRefName: string
}
