# Security policy

## Supported versions

Security fixes target the latest release. Older releases may not receive backports.

## Reporting a vulnerability

Please do not disclose exploitable details in a public issue. Use the repository's **Security → Advisories → Report a vulnerability** form. Maintainers must enable GitHub private vulnerability reporting before making the repository public. If that form is unavailable, contact a maintainer privately through a contact method on their GitHub profile.

Include the affected version, a minimal reproduction using synthetic files, impact, and any suggested mitigation. We will investigate privately and coordinate a fix and advisory before public disclosure where appropriate.

## Security model

Office MCP runs locally with the user's filesystem and Office permissions. It is not a network sandbox. Limit `OFFICE_MCP_ROOTS` to the folders the agent actually needs, review write operations, and do not expose the stdio process to untrusted users. Native Office automation may execute Office behaviors associated with opening a file; do not process untrusted Office files outside a suitably isolated environment. The quality gate checks package structure and explicit assertions, not document safety or truthfulness.
