# Akamai CPS — FQDN to Enrollment ID lookup

## Purpose

Allow the existing certificate scheduling workflow to accept an FQDN, trim surrounding whitespace, and perform a case-insensitive **exact** match against certificate Common Name (CN) and Subject Alternative Names (SANs). Preserve the existing `--enrollment-id` workflow.

**Files:** `akamai_cert_manager_proxy.py` and `akamai-cert-schedule-deploy.groovy`.

> **Production caution:** The lookup below fetches every enrollment individually. The existing `get_enrollment()` also writes an output JSON file for each call. This can generate significant API traffic and hit CPS rate limits. Verify the enrollment count and test lookup separately before running a production scheduling job. The existing `get_enrollments()` returns `[]` on API failure, so an empty list is treated as failure/no result; it will not schedule anything.

## 1. Add a method to `AkamaiCertManager`

In `akamai_cert_manager_proxy.py`, insert this method **inside** `class AkamaiCertManager`, immediately after the existing `get_enrollment()` method and before `get_deployments()` (approximately line 618). Do **not** add the alternative direct-API implementation.

```python
    # --------------------------------------------------
    # Resolve Enrollment ID by FQDN
    # --------------------------------------------------
    def get_enrollment_id_by_fqdn(self, fqdn):
        """Find a unique CPS enrollment by case-insensitive exact CN/SAN match."""

        target = (fqdn or "").strip().casefold()

        if not target:
            log("ERROR", "FQDN cannot be empty")
            return None

        log("INFO", f"Searching CPS enrollments for FQDN: {target}")
        enrollments = self.get_enrollments()

        if not enrollments:
            log("ERROR", "No CPS enrollments returned (or list request failed)")
            return None

        matches = []
        failed_lookups = []

        for item in enrollments:
            enrollment_id = item.get("id")
            if enrollment_id is None:
                failed_lookups.append("missing enrollment ID in list response")
                continue

            enrollment = self.get_enrollment(enrollment_id)
            if not enrollment:
                failed_lookups.append(enrollment_id)
                continue

            csr = enrollment.get("csr") or {}
            cn = (csr.get("cn") or "").strip().casefold()
            sans = [
                san.strip().casefold()
                for san in (csr.get("sans") or [])
                if isinstance(san, str)
            ]

            if target == cn or target in sans:
                matches.append(enrollment_id)

        matches = list(dict.fromkeys(matches))

        # Never trust a partial lookup: a failed detail request could hide a match.
        if failed_lookups:
            raise RuntimeError(
                "Incomplete CPS FQDN lookup; failed enrollment lookups: "
                f"{failed_lookups}"
            )

        if len(matches) == 1:
            log("INFO", f"Resolved FQDN '{target}' to Enrollment ID {matches[0]}")
            return matches[0]

        if len(matches) > 1:
            raise ValueError(
                f"Multiple CPS enrollments matched FQDN '{target}': {matches}"
            )

        log("ERROR", f"No CPS enrollment matched FQDN '{target}'")
        return None
```

## 2. Keep `--fqdn` and normalise its value

In `main()`, retain your existing parser argument:

```python
parser.add_argument(
    "--fqdn",
    type=str,
    help="Certificate FQDN used to resolve the CPS Enrollment ID"
)
```

Immediately **after** `args = parser.parse_args()`, add:

```python
if args.fqdn:
    args.fqdn = args.fqdn.strip().casefold()
```

## 3. Resolve the Enrollment ID before scheduling

Find the existing scheduling branch (approximately line 1450):

```python
if args.report_only and args.schedule:
```

Insert this block **inside that branch**, before the existing `if not args.enrollment_id:` validation:

```python
    # Resolve Enrollment ID from FQDN only when an ID was not supplied.
    if args.fqdn and not args.enrollment_id:
        log("INFO", f"Resolving Enrollment ID for FQDN: {args.fqdn}")

        try:
            args.enrollment_id = manager.get_enrollment_id_by_fqdn(
                args.fqdn
            )
        except (ValueError, RuntimeError) as e:
            log("ERROR", f"FQDN lookup failed: {e}")
            return

        if not args.enrollment_id:
            log("ERROR", f"Unable to resolve Enrollment ID for FQDN: {args.fqdn}")
            return
```

Keep the rest of the scheduling block intact. Optionally update the existing missing-ID error text to say `Provide --enrollment-id or --fqdn`.

**Note:** In this implementation a lookup failure returns from `main()`. If Jenkins needs a nonzero exit code to mark the stage failed, replace the failure-path `return` statements with `raise RuntimeError(...)` (and allow it to propagate). Do not leave a failed lookup reporting a successful job unintentionally.

## 4. Update Jenkins Groovy

In `akamai-cert-schedule-deploy.groovy`, immediately before the `sh` invocation, validate the FQDN parameter. This example assumes the existing command is inside a Groovy scope where `params` and `REGION` are available.

```groovy
def fqdnValue = (params.FQDN ?: '').trim()

if (!(fqdnValue ==~ /(?i)^(?=.{1,253}$)[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)*$/)) {
    error('Invalid FQDN supplied')
}
```

Then replace the original command:

```groovy
python akamai_cert_manager_proxy.py --section '${REGION}' --report-only --schedule --fqdn ${params.FQDN}
```

with the command below **inside the existing interpolated Groovy `sh` string**:

```groovy
python akamai_cert_manager_proxy.py --section '${REGION}' --report-only --schedule --fqdn '${fqdnValue}'
```

The hostname validation prevents quote and shell-metacharacter injection via `fqdnValue`. It accepts conventional DNS hostnames, **not** wildcard names such as `*.example.com`. If wildcard certificates need to be searched, adjust the validation and matching rules deliberately.

## 5. Test before production scheduling

1. Verify that `get_enrollments()` returns a list of dictionaries with `id` and that `get_enrollment()` returns `csr.cn` and `csr.sans`.
2. Test a known FQDN using uppercase letters and surrounding spaces; expect exactly one Enrollment ID.
3. Test an unknown FQDN; confirm no scheduling takes place.
4. Test duplicate CN/SAN matches across multiple enrollments; confirm lookup stops rather than choosing arbitrarily.
5. Test the existing `--enrollment-id` path; confirm it remains unchanged.
6. Check API traffic, rate limiting and generated JSON files before enabling routine production use.
7. Ensure lookup failures result in a **failed Jenkins stage** rather than a success exit code.

## Expected behaviour

| Input | Result |
|---|---|
| `API.HSBC.COM` | Matches `api.hsbc.com` |
| ` api.hsbc.com ` | Matches `api.hsbc.com` |
| FQDN appears in SANs | Resolves the matching Enrollment ID |
| Unknown FQDN | No match; do not schedule |
| Multiple matching enrollments | Error; do not schedule |
| Existing `--enrollment-id 12345` | Existing scheduling path unchanged |

**Scope:** This change only resolves the Enrollment ID. It does not change renewal, deployment scheduling, polling or notifications.
