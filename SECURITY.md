# Security Policy

## Reporting Security Issues

**Do not open public GitHub issues for security vulnerabilities.**

Please report security issues to Microsoft's Security Response Center (MSRC) at [https://msrc.microsoft.com/create-report](https://msrc.microsoft.com/create-report) or email [secure@microsoft.com](mailto:secure@microsoft.com).

## Security Considerations

When using this script:

1. **Credentials** — Never hardcode credentials. Use `Connect-AzAccount` for interactive authentication.
2. **Output Files** — Excel files contain sensitive network information. Store securely.
3. **Dependencies** — Keep modules updated: `Update-Module Az.Accounts, Az.ResourceGraph, ImportExcel -Force`
4. **No Secrets** — Never commit credentials, subscription IDs, or secrets to the repo.

## Supported Versions

| Version | Status |
|---------|--------|
| 1.0+ | Supported |

Security updates will be released as soon as possible after discovery.
