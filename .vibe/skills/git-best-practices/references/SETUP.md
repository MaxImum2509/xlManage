# Repository Setup

## Local Creation

### New project

```bash
# 1. Create folder
mkdir my-project
cd my-project

# 2. Initialize Git
git init

# 3. Create .gitignore
curl -o .gitignore https://raw.githubusercontent.com/github/gitignore/main/Python.gitignore

# 4. First commit
git add .
git commit -m "chore: initial commit"
```

### Existing project

```bash
cd existing-project
git init
git add .
git commit -m "chore: initial commit"
```

## Git Configuration

### Global (once)

```bash
git config --global user.name "Your Name"
git config --global user.email "email@example.com"
git config --global init.defaultBranch main
git config --global pull.rebase true
```

### Per project (optional)

```bash
git config user.email "work@company.com"
```

## GitHub Connection

### HTTPS method (recommended for beginners)

```bash
# Push to GitHub
git remote add origin https://github.com/username/repo.git
git branch -M main
git push -u origin main
```

### SSH method (recommended)

```bash
# 1. Generate SSH key
ssh-keygen -t ed25519 -C "email@example.com"

# 2. Copy public key
cat ~/.ssh/id_ed25519.pub
# Paste in GitHub → Settings → SSH Keys

# 3. Connect
git remote add origin git@github.com:username/repo.git
git push -u origin main
```

## Creation on GitHub

### Via CLI (GitHub CLI)

```bash
# Install: https://cli.github.com/
gh auth login
gh repo create my-project --public --source=. --push
```

### Via Web Interface

1. GitHub → New Repository
2. Do NOT initialize (README, .gitignore)
3. Copy the "push existing repository" commands

## Recommended Structure

```
my-project/
├── .git/              # Git internal
├── .gitignore         # Ignored files
├── README.md          # Documentation
├── LICENSE            # License
└── src/               # Source code
```

## Verification

```bash
git status          # Working directory status
git log --oneline   # Commit history
git remote -v       # Configured remotes
```

## Useful Aliases

```bash
git config --global alias.st "status -sb"
git config --global alias.lg "log --oneline --graph --decorate"
```
