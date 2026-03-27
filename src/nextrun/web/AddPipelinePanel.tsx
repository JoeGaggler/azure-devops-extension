import { useState } from 'react'

import { TitleSize } from "azure-devops-ui/Header";
import { Panel } from "azure-devops-ui/Panel";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";

export function AddPipelinePanel(p: AddPipelinePanelProps) {
    const [id, setId] = useState("")
    const [name, setName] = useState("")
    const [resource, setResource] = useState("")

    async function onCreate() {
        let idNum = Number(id);
        if (isNaN(idNum)) {
            // TODO: show error
            return;
        }

        let data: AddPipelinePanelValues = {
            id: idNum,
            name: name,
            resource: resource,
        }
        await p.onCommit(data);
    }

    async function onCancel() {
        await p.onCancel();
    }

    return (
        <Panel
            onDismiss={() => p.onCancel()}
            titleProps={
                {
                    text: "New Pipeline",
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
            ]}
        >
            <div className="flex-column rhythm-vertical-8">
                <TextField
                    label={"Id"}
                    value={id}
                    onChange={(e, nextValue) => e && setId(nextValue)}
                    width={TextFieldWidth.standard}
                />

                <TextField
                    label={"Name"}
                    value={name}
                    onChange={(e, nextValue) => e && setName(nextValue)}
                    width={TextFieldWidth.standard}
                />

                <TextField
                    label={"Resource"}
                    value={resource}
                    onChange={(e, nextValue) => e && setResource(nextValue)}
                    width={TextFieldWidth.standard}
                />
            </div>
        </Panel>
    )
}

export interface AddPipelinePanelValues {
    id: number
    name: string
    resource: string
}

export interface AddPipelinePanelProps {
    onCommit: (data: AddPipelinePanelValues) => Promise<void>
    onCancel: () => Promise<void>
}
