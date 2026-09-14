# Keep Plans Gitignored

Apply this rule whenever creating, updating, moving, or managing implementation plans, design documents, or task roadmaps in `docs/plans/` (including `docs/plans/completed/` and any other plan subdirectories).

## Rules

1. **Plans remain gitignored:**
   All files in `docs/plans/` and its subdirectories must remain gitignored. Plans are local working documents and agent scratchpads, not tracked repository artifacts.

2. **Never create `.gitignore` exceptions for plans:**
   Do not add whitelist exceptions (e.g., `!docs/plans/<filename>.md` or `!docs/plans/completed/<filename>.md`) to `.gitignore`.

3. **Never stage or commit plans:**
   Do not add (`git add`), stage, or commit any plan files to version control unless the user explicitly and directly requests that a specific plan file be committed.
