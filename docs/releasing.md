# How to Release

All PRs must be merged into `main` before tagging.

## Steps

```bash
git checkout main
git pull origin main
git tag v1.0.0
git push origin v1.0.0
```

GitHub Actions will automatically:
1. Build `PDF-to-Excel-Converter.app` on macOS
2. Build `PDF-to-Excel-Converter.exe` on Windows
3. Publish a GitHub Release with both files at:
   `https://github.com/MoCha15/pdf-to-excel-converter/releases`

## Versioning

| Change type | Example |
|---|---|
| Bug fix | `v1.0.1` |
| New feature | `v1.1.0` |
| Breaking change | `v2.0.0` |
| Pre-release / beta | `v1.1.0-beta.1` ← auto-marked as pre-release |
