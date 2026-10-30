<!-- BEGIN:emeltec-team-rules (synced from EmelTI2026/Eme_vault team-rules/AGENTS.team.md; edit it there) -->

## Emeltec team rules

These rules apply to every AI agent working in this repo. Project-specific instructions after this block
may add to them; where they conflict, the project-specific instruction wins.

### Commits and git

- Use Conventional Commits (`feat:`, `fix:`, `docs:`, `chore:`, `ci:`, `refactor:`, `test:`).
- Never add `Co-Authored-By` lines or any AI attribution to commits.
- Never push directly to `main`: work on a branch and open a pull request.
- Never merge a pull request without the user's explicit approval.

### Language

- Reply to the user in neutral Spanish (no voseo), unless the user writes in another language.
- Commit messages, code comments and docs follow the language this repo already uses (check `git log`
  and the existing docs). A new repo defaults to English.

### Engram memory

- When Engram is available, save decisions, bug fixes, discoveries and conventions with `mem_save`
  proactively, without waiting to be asked.
- The memory project comes from `.engram/config.json`; never save under another project name.
- Search memory (`mem_search`) before redoing or re-deciding past work.

### How to reply

- Keep answers short; expand only when the user asks or the task needs it.
- Ask at most one question at a time, then stop and wait for the answer.
- Verify technical claims in the code or docs before stating them; say you will verify first.
- If the user is wrong, explain why with evidence; if you were wrong, say so and show the proof.
- Offer alternatives with their tradeoffs only when there is a real choice to make.
<!-- END:emeltec-team-rules -->
