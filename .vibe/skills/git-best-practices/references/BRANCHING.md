# Branching Strategies

## Workflows

### GitHub Flow (Recommended for solo)

```
main
 ↑
  └─ feature/name → merge
```

**Process**:
1. `git checkout -b feature/name`
2. Commits
3. `git checkout main && git merge feature/name`
4. `git branch -d feature/name`

### Git Flow (Complex teams)

```
main ← release ← develop ← feature
```

Branches: `feature/`, `release/`, `hotfix/`

## Naming

| Pattern | Usage |
|---------|-------|
| `feature/name` | New feature |
| `bugfix/name` | Bug fix |
| `hotfix/name` | Critical fix |
| `refactor/name` | Refactoring |
| `docs/name` | Documentation |

**Rules**: kebab-case, descriptive, no special characters

## Lifespan

| Type | Duration |
|------|----------|
| Feature | 1-3 days |
| Bugfix | < 1 day |
| Hotfix | A few hours |

## Best Practices

✅ **Do**:
- Sync regularly: `git rebase main`
- Delete after merge
- Atomic commits

❌ **Avoid**:
- Orphan branches
- Commits on main (GitHub Flow workflow)
- Branches > 1 week

## Aliases

```bash
alias gnb="git checkout -b feature/"
alias gsync='git fetch && git rebase origin/main'
alias gfinish='git checkout main && git merge @{-1} && git branch -d @{-1}'
```
