function issueKey(issue) {
    return `${issue?.source || ''}\u001f${issue?.severity || ''}\u001f${issue?.code || ''}\u001f${issue?.message || ''}`;
}

/** Subtracts baseline validation issues as a multiset, preserving duplicates. */
export function subtractValidationIssueMultiset(outputIssues = [], baselineIssues = []) {
    const remaining = new Map();
    for (const issue of baselineIssues || []) {
        const key = issueKey(issue);
        remaining.set(key, (remaining.get(key) || 0) + 1);
    }
    return (outputIssues || []).filter(issue => {
        const key = issueKey(issue);
        const count = remaining.get(key) || 0;
        if (count === 0) return true;
        remaining.set(key, count - 1);
        return false;
    });
}

export function validationErrors(issues = []) {
    return (issues || []).filter(issue => issue?.severity === 'error');
}
