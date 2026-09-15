# Chess Player Verifier — Cloudflare Python Worker

Cloudflare-ready version of the local chess player verifier.

Features:
- Excel upload
- AICF or UPCSA mode
- Exact AICF ID matching
- UPCSA player-page lookup
- Active / Not Active membership status and expiry where available
- Duplicate-ID caching per workbook
- Browser-like UPCSA headers and retry policy

## Deployment

This is a **Cloudflare Worker**, not a static Cloudflare Pages site.

Cloudflare's current Python Worker tooling uses `pywrangler`. If deploying locally:

```bash
uv run pywrangler dev
uv run pywrangler deploy
```

You can also connect the GitHub repository through Cloudflare's Workers/Git integration.

## Important

Python Workers are currently in open beta. Test this with the workbook sizes you expect before relying on it for large batches.

UPCSA requests intentionally wait between requests and retry failures.
