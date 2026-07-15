export function cutPrefix(str: string, prefix: string): string {
    return str.startsWith(prefix) ? str.substring(prefix.length) : str;
}

export function distinctBy<T, K>(array: T[], keySelector: (item: T) => K): T[] {
    const seen = new Set<K>();
    return array.filter(item => {
        const key = keySelector(item);
        if (seen.has(key)) {
            return false;
        }
        seen.add(key);
        return true;
    });
}

export function firstDefined<T>(...values: (T | undefined)[]): T | undefined {
    for (const value of values) {
        if (value !== undefined) {
            return value;
        }
    }
    return undefined;
}