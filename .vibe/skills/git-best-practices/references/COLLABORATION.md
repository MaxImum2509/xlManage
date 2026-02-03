# Collaboration and Pull Requests

## Pull Requests

### Structure

```markdown
## Description
[Summary]

## Context
[Why this change?]

## Testing
[How to verify]

Closes #123
```

### Rules

- Title: commit convention
- Description: explains the "why"
- Size: < 400 lines
- CI passing

## Code Review

### For the reviewer

1. Understand the context (description, issues)
2. Verify the approach
3. Review: architecture, complexity, tests
4. Constructive comments with alternatives

### Comment template

```markdown
**Issue**: [description]
**Suggestion**: [solution]
**Example**: [code]
```

## Conflicts

```bash
git checkout feature/name
git pull origin main
# Resolve conflicts
git add .
git commit -m "resolve: merge conflicts"
```

**Prevention**:
- Sync regularly
- Short branches
- Frequent commits

## Merge

| Strategy | Usage |
|----------|-------|
| **Squash** | Solo, small PRs |
| **Merge commit** | Team, important history |

## Checklist

**Author**:
- [ ] CI passing
- [ ] Complete description
- [ ] Tests added
- [ ] No debug code

**Reviewer**:
- [ ] Architecture OK
- [ ] Code readable
- [ ] Tests sufficient
