# Security policy

## Supported versions

Security fixes go into the latest release of the newest minor version. Until 1.0 that is 0.8.x; from 1.0, the latest 1.x.

| Version | Supported |
|---|---|
| 0.8.x | ✅ |
| < 0.8 | ❌ (upgrade to 0.8.x; the [changelog](CHANGELOG.md) lists every change) |

## Reporting a vulnerability

Please don't open a public issue. Report it privately through GitHub: on the repository's **Security** tab, choose **Report a vulnerability** ([direct link](https://github.com/7a6163/fast_xlsx/security/advisories/new)).

Include what an attacker could do, the fast_xlsx, Ruby and platform versions, and a script or input that shows it, if you have one.

What to expect:

- An acknowledgement within a week.
- A fix in a patch release, and a GitHub security advisory crediting you (unless you'd rather not be named), once users can upgrade.
- If it turns out to be in a dependency (rust_xlsxwriter, magnus, rb-sys, the Rust zip crates), a report to that project as well.

## Scope

fast_xlsx only writes `.xlsx` files: it doesn't read or parse spreadsheets. Issues of interest include memory-safety bugs in the native extension (a crash or corruption reachable from Ruby), and output that makes Excel or other readers misbehave from data an application would reasonably pass in (cell text, formulas, URLs, images). Formulas and hyperlinks are written as given: an application writing untrusted input as a `FastXlsx::Formula` or `FastXlsx::URL` should validate it first, as with CSV formula injection.
