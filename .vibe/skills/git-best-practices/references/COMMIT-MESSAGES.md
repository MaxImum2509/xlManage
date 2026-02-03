# Commit Messages

## Règle d'Or : EN ANGLAIS UNIQUEMENT

**TOUS les messages de commit doivent être rédigés en anglais**, quelle que soit la langue de l'utilisateur, de l'équipe ou du projet. Cette règle est **non négociable**.

```
✅ feat(auth): add OAuth support
❌ feat(auth): ajoute support OAuth
```

## Format

```
<type>(<scope>): <subject>

<body>

<footer>
```

### Rules

| Section | Rules |
|---------|-------|
| **Subject** | < 50 chars, imperative, lowercase, no final period |
| **Body** | < 72 chars/line, explains the **why** |
| **Footer** | `Closes #123`, `BREAKING CHANGE: ...` |

## Types

| Type | Usage |
|------|-------|
| `feat` | New feature |
| `fix` | Bug fix |
| `docs` | Documentation |
| `refactor` | Code refactoring |
| `test` | Tests |
| `chore` | Maintenance |
| `style` | Formatting |
| `perf` | Performance |
| `build` | Build system |
| `ci` | CI/CD |

## Example

```
feat(auth): add Google OAuth support

Allow users to sign in with their Google account.
Uses standard OAuth 2.0 library.

Closes #42
```

## Anti-patterns

```
❌ fix              # Too vague
❌ update           # No context
❌ Added feature    # Past tense, not imperative
❌ feat(auth): ajoute support OAuth  # French - NOT ALLOWED
✅ feat(auth): add OAuth support     # English - CORRECT
```

## Checklist

- [ ] **EN ANGLAIS** - Quelle que soit la langue de l'utilisateur
- [ ] Subject < 50 chars, imperative, lowercase
- [ ] Body explains the "why"
- [ ] No final period in subject
- [ ] Body < 72 chars/line
- [ ] One logical change per commit
