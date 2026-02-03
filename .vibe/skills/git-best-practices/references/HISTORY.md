# History and Cleanup

## Golden Rule

Manipulate history **ONLY** on local branches.

## Commands

### Amend (last local commit)

```bash
git add forgotten_file.txt
git commit --amend
```

⚠️ Never after push

### Interactive rebase

```bash
git rebase -i HEAD~3
```

**Commands**:
- `pick`: keep
- `reword`: edit message
- `squash`: merge
- `drop`: delete

### Squash

```bash
git merge --squash feature/name
git commit -m "feat: complete feature"
```

### Reset

| Option | Effect |
|--------|--------|
| `--soft` | Undo commit, keep staged |
| `--mixed` | Undo commit, unstaged (default) |
| `--hard` | Destroys everything ⚠️ |

```bash
git reset --soft HEAD~1  # Undo last commit
git reset --hard HEAD~3  # ⚠️ Destructive
```

### Revert (safe)

```bash
git revert abc1234  # Creates inverse commit
```

✅ Always on shared branches

### Cherry-pick

```bash
git cherry-pick abc1234  # Copy commit
```

## Cleanup

```bash
git branch -d $(git branch --merged)  # Merged branches
git clean -fd                         # Untracked files
git gc                                # Garbage collection
```

## Rebase vs Merge

| Operation | Usage | Danger |
|-----------|-------|--------|
| `amend` | Last local commit | ⚠️ After push |
| `rebase` | Clean history | ⚠️ After push |
| `squash` | Merge commits | ⚠️ After push |
| `reset --hard` | Go back | 🔴 Destructive |
| `revert` | Undo commit | ✅ Safe |
| `push --force` | Force push | ⚠️ ⚠️ ⚠️ |

**Safe force push**: `git push --force-with-lease`

## Tools

```bash
git bisect start      # Find bug by dichotomy
git reflog            # Reference history
```
