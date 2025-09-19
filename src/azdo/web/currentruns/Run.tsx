import { RunIcon, StatusType } from "./RunStatus";
import * as luxon from 'luxon';
import { PillGroup } from "azure-devops-ui/PillGroup";
import { Pill, PillVariant, PillSize } from "azure-devops-ui/Pill";

interface RunProps {
    name: string;
    definitionName: string;
    status: StatusType;
    comment: string;
    started: number | null;
    isAlternate: boolean;
    isKnown: boolean;
    knownTags: string[];
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
                {
                    <div className="flex-row flex-grow">
                        <div className="flex-grow" />
                        {
                            p.isKnown && (
                                <PillGroup className="padding-left-16 padding-right-16">
                                    {p.knownTags.map((tag: string) => (
                                        <Pill key={tag} size={PillSize.compact} variant={PillVariant.outlined} color={{ red: 64, green: 128, blue: 64 }}>{tag}</Pill>
                                    ))}
                                </PillGroup>
                            )
                        }
                        {p.started && (<div>{luxon.DateTime.fromSeconds(p.started).toRelative()}</div>)}
                    </div>
                }
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
