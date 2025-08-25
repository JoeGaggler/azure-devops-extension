import { RunIcon, StatusType } from "./RunStatus";
import * as luxon from 'luxon';


interface RunProps {
    name: string;
    definitionName: string;
    status: StatusType;
    comment: string;
    started: number | null;
    isAlternate: boolean;
    isKnown: boolean;
}

function Run(p: RunProps) {
    let extra = "";
    if (p.isAlternate) { extra += " alternate-run-row"; } else { extra += " normal-run-row"; }
    let className = `scroll-hidden flex-row flex-center flex-grow padding-4 ${extra}`;
    return (
        <>
            <div className={className}>
                <div className="margin-right-4"></div>
                <RunIcon status={p.status} className="nothing_here" />
                <div className="font-size-m flex-self-center padding-4 flex-noshrink">{p.name}</div>
                <div className="font-size-ml flex-self-center padding-4 flex-noshrink">·</div>
                <div className="font-size-m flex-self-center padding-4 flex-noshrink italic">{p.definitionName}</div>
                <div className="font-size-ml flex-self-center padding-4 flex-noshrink">·</div>
                {p.comment && (<div className="text-ellipsis text-neutral-70 padding-left-8">{p.comment}</div>)}
                {p.started && (<div className="flex-row flex-grow"><div className="flex-grow" />{!p.isKnown && (<div className="padding-right-8">·</div>)}<div>{luxon.DateTime.fromSeconds(p.started).toRelative()}</div></div>)}
            </div>

        </>
    )
}

function GetRunStatusType(state: string | undefined, result: string | undefined): StatusType {
    if (state == "notStarted" || state == "postponed") {
        return "Waiting";
    }
    if (state == "inProgress") {
        return "Running";
    }
    if (state === "completed") {
        if (result === "succeeded") {
            return "Success";
        }
        if (result === "skipped") {
            return "Skipped";
        }
        if (result == "partiallySucceeded") {
            return "Warning";
        }
        if (result == "succeededWithIssues") {
            return "Warning";
        }
        if (result === "failed") {
            return "Failed";
        }
        if (result == "canceled") {
            return "Canceled";
        }
    }
    return "Queued";
}

export { Run, GetRunStatusType };
export type { RunProps, StatusType };
