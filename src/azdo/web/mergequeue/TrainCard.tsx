import React from "react";
import { Card } from "azure-devops-ui/Card";
import { Icon } from "azure-devops-ui/Icon";

interface TrainCardProps {
    name: string;
}

function TrainCard(p: TrainCardProps) {
    const [collapsed, setCollapsed] = React.useState<boolean>(true);

    function onCollapseClicked() {
        setCollapsed(!collapsed);
    }

    return (
        <Card
            className="flex-grow"
            collapsible={true}
            collapsed={collapsed}
            onCollapseClick={() => onCollapseClicked()}
            titleProps={{ text: p.name, ariaLevel: 2 }}
        >
            <div className="flex-row" style={{ flexWrap: "wrap" }}>
                <Icon iconName="Video" />
            </div>
        </Card>
    );
}

export { TrainCard };
export type { TrainCardProps };