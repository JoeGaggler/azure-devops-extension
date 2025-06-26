import { Card } from "azure-devops-ui/Card";

interface AppProps {
    bearerToken: string | null;
    appToken: string | null;
}

function App(p: AppProps) {
    if (p) {
        console.log("App props:", p);
    }

    return (
        <>
            <Card className="padding-8">
                <div className="flex-row rhythm-horizontal-8 flex-center">
                    <div>Coming soon.</div>
                </div>
            </Card>
        </>
    )
}

export { App };
export type { AppProps };
