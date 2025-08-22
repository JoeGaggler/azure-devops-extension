import { Status, Statuses, StatusSize, StatusType } from "azure-devops-ui/Status";

function RunIcon({ status, className }: { status: StatusType, className: string }) {
    return (
        <>
            {
                status === "Running" ? (
                    <Status
                        {...Statuses.Running}
                        key="running"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                status === "Failed" ? (
                    <Status
                        {...Statuses.Failed}
                        key="failed"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                status === "Warning" ? (
                    <Status
                        {...Statuses.Warning}
                        key="warning"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                // blue i
                status === "Information" ? (
                    <Status
                        {...Statuses.Information}
                        key="information"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                status === "Success" ? (
                    <Status
                        {...Statuses.Success}
                        key="success"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                // blue clock
                status === "Waiting" ? (
                    <Status
                        {...Statuses.Waiting}
                        key="waiting"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                // white circle
                status === "Queued" ? (
                    <Status
                        {...Statuses.Queued}
                        key="queued"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                // white slash
                status === "Canceled" ? (
                    <Status
                        {...Statuses.Canceled}
                        key="canceled"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
            {
                // white chevron
                status === "Skipped" ? (
                    <Status
                        {...Statuses.Skipped}
                        key="skipped"
                        size={StatusSize.m}
                        className={className}
                    />
                )
                    : <></>
            }
        </>
    )
}

export { RunIcon };
export type { StatusType };
