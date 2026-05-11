# Publish to GitHub

## Option A: GitHub CLI

Install and authenticate GitHub CLI, then run from this repository folder:

```bash
gh auth login
git init
git add .
git commit -m "Initial release of EndNote Field Code Converter"
gh repo create endnote-fieldcode-converter \
  --public \
  --source=. \
  --remote=origin \
  --push \
  --description "ChatGPT Skill and Python script for converting plain-text numeric citations in DOCX manuscripts into EndNote field codes."
```

## Option B: Existing empty repository

Create an empty public repository on GitHub, then run:

```bash
git init
git add .
git commit -m "Initial release of EndNote Field Code Converter"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/endnote-fieldcode-converter.git
git push -u origin main
```

## Suggested GitHub topics

```text
endnote docx word references citations manuscript biomedical-writing chatgpt-skill ooxml crossref
```

## License

The repository includes an MIT License in `LICENSE`. Keep this file at the repository root so GitHub can detect the license automatically.
