---
name: git-best-practices
description: Concise Git best practices guide for solo developers and small teams. Use when asking for Git advice, reviewing workflows, improving Git practices, setting up Git conventions, or troubleshooting Git operations.
---

## Quick Reference

| Topic | See |
|-------|-----|
| **Setup** | [references/SETUP.md](references/SETUP.md) |
| **Commit Messages** | [references/COMMIT-MESSAGES.md](references/COMMIT-MESSAGES.md) |
| **Branching Strategies** | [references/BRANCHING.md](references/BRANCHING.md) |
| **Collaboration & PRs** | [references/COLLABORATION.md](references/COLLABORATION.md) |
| **History & Cleanup** | [references/HISTORY.md](references/HISTORY.md) |
| **Commit Template** | [assets/commit-template.txt](assets/commit-template.txt) |

## Fundamental Rule: English Commits

**ALL commit messages MUST be in English**, regardless of the user's or team's language. This is a non-negotiable rule to maintain a universally readable history.

```
✅ feat(auth): add OAuth support
❌ feat(auth): ajoute support OAuth
```

## Principles

1. **Atomic commits**: One commit = one logical unit, message explaining the **why**
2. **Clean history**: Interactive rebase on local branches before push
3. **Short branches**: 1-3 days maximum to minimize conflicts
4. **Safety first**: Manipulate history only on local branches

## Workflow

### Solo Developer

```bash
git checkout -b feature/name
git commit -m "feat: description"
git checkout main
git merge feature/name --squash
git commit -m "feat: full description"
git branch -d feature/name
```

### Team with PR

```bash
git checkout -b feature/name
git commit -m "feat: description"
git push origin feature/name
# Create PR with full description
# Review and merge
```

## Essential Rules

### Commit Messages (ENGLISH ONLY)
- **Language**: **ENGLISH mandatory**, regardless of context
- Types: `feat`, `fix`, `docs`, `refactor`, `test`, `chore`
- Subject: < 50 characters, imperative, lowercase (except type), no final period
- Body: explains the "why", < 72 characters/line
- Footer: `Closes #123`

### Branching
- `main`: production, always deployable
- `feature/name`: kebab-case, descriptive
- Duration: < 3 days
- Delete after merge

### History
| Operation | Local | Shared |
|-----------|-------|--------|
| `git commit --amend` | ✅ | ❌ |
| `git rebase -i` | ✅ | ⚠️ |
| `git reset --hard` | ⚠️ | 🔴 |
| `git revert` | ✅ | ✅ |
| `git push --force` | ⚠️ | ❌ |

## Useful Aliases

```bash
git config --global alias.undo "reset --soft HEAD~1"
git config --global alias.rbi "rebase -i HEAD~3"
git config --global alias.lg "log --oneline --graph --decorate --all"
```

## Anti-patterns

```
❌ fix              # Too vague
❌ Added feature    # Past tense
❌ feat(auth): ajoute support OAuth  # French!
✅ feat(auth): add OAuth support     # English
```
