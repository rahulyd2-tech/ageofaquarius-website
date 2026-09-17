# ageofaquariusdevelopers.in

Static company website for Age of Aquarius Group (single `index.html`, no build step), served by GitHub Pages.

## GitHub Pages
Settings → Pages → Source: **Deploy from a branch**, Branch: **main** / **(root)**.
Custom domain: `ageofaquariusdevelopers.in` (kept in `CNAME`). Tick **Enforce HTTPS** once the certificate is issued.

## DNS (at the domain registrar)
| Type | Host | Value |
|---|---|---|
| A | @ | 185.199.108.153 |
| A | @ | 185.199.109.153 |
| A | @ | 185.199.110.153 |
| A | @ | 185.199.111.153 |
| AAAA | @ | 2606:50c0:8000::153 |
| AAAA | @ | 2606:50c0:8001::153 |
| AAAA | @ | 2606:50c0:8002::153 |
| AAAA | @ | 2606:50c0:8003::153 |
| CNAME | www | rahulyd2-tech.github.io |

Remove any other A/AAAA/CNAME records for `@` and `www` (e.g. registrar parking).
