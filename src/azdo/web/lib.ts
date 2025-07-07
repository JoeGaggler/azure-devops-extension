import { ListSelection } from "azure-devops-ui/List";

export function sortByString<T>(array: T[], f: (i: T) => string) {
    array.sort((a, b) => {
        let x = f(a);
        let y = f(b);
        if (x === y) {
            return 0;
        }
        if (x === undefined || x === null) {
            return 1;
        }
        if (y === undefined || y === null) {
            return -1;
        }
        return x.localeCompare(y);
    });
}

export function sortByNumber<T>(array: T[], f: (i: T) => number) {
    array.sort((a, b) => {
        let x = f(a);
        let y = f(b);
        if (x === y) {
            return 0;
        }
        if (x === undefined || x === null) {
            return 1;
        }
        if (y === undefined || y === null) {
            return -1;
        }
        return x - y;
    });
}

export function applySelections<T, TItem>(listSelection: ListSelection, all: T[], accessor: (t: T) => TItem, ids: TItem[]) {
    listSelection.clear();
    for (let id of ids) {
        let idx = all.findIndex((item) => accessor(item) === id);
        if (idx < 0) { continue; }
        listSelection.select(idx, 1, true, true);
    }
}
