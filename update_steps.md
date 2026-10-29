# Akamai CPS FQDN lookup-only steps

## 1. Add CLI argument

After `--fqdn` and before `parse_args()`:

```python
parser.add_argument("--lookup-only", action="store_true", help="Resolve FQDN to CPS Enrollment ID and exit")
```

## 2. Add early-exit block

After `manager.auto_poll = args.auto_poll`, before StateTracker initialization:

```python
    if args.lookup_only:
        if not args.fqdn or args.schedule or args.renew or args.get_change_status or args.cert_pinning:
            log("ERROR", "Lookup-only requires --fqdn and cannot be combined with change operations")
            raise SystemExit(1)
        try:
            enrollment_id = manager.get_enrollment_id_by_fqdn(args.fqdn)
            if not enrollment_id:
                raise RuntimeError("FQDN not found")
            if args.enrollment_id and args.enrollment_id != enrollment_id:
                raise ValueError("Enrollment ID does not match FQDN")
            log("SUCCESS", f"FQDN {args.fqdn} resolved to Enrollment ID {enrollment_id}")
            return
        except (ValueError, RuntimeError) as exc:
            log("ERROR", f"FQDN lookup failed: {exc}")
            raise SystemExit(1)
```

Use four spaces for each indentation level. The resolver must use `get_enrollment(..., save_json=False)`.

## 3. Compile

```bash
python -m py_compile akamai_cert_manager_proxy.py
```

## 4. DEV positive test

```bash
python akamai_cert_manager_proxy.py --section DEV --fqdn cyber-wasp-mismatch-test.hsbc.com.hk --lookup-only
echo $?
```

Expected Enrollment ID: **333961**, exit code **0**.

## 5. GLOBAL positive test

```bash
python akamai_cert_manager_proxy.py --section GLOBAL --fqdn www.epmerchant.com.hk --lookup-only
echo $?
```

Expected Enrollment ID: **271148**, exit code **0**.

## 6. Negative tests

```bash
python akamai_cert_manager_proxy.py --section DEV --fqdn nonexistent-test.invalid --lookup-only
echo $?
python akamai_cert_manager_proxy.py --section DEV --lookup-only
echo $?
python akamai_cert_manager_proxy.py --section DEV --fqdn cyber-wasp-mismatch-test.hsbc.com.hk --lookup-only --schedule
echo $?
python akamai_cert_manager_proxy.py --section DEV --fqdn cyber-wasp-mismatch-test.hsbc.com.hk --enrollment-id 999999 --lookup-only
echo $?
```

Each must fail with exit code **1**. The schedule conflict must be rejected before any scheduling call.

## 7. Review

```bash
git status --short
git diff --check
git diff -- akamai_cert_manager_proxy.py
```

Confirm no certificate JSON files, Git operations, renewal, deployment scheduling, pinning, or change acknowledgements occurred. Check that `setup_session()` and lookup helpers are read-only. Large-account lookup performance and pagination completeness remain to be verified before Jenkins rollout. Do not use `--report-only --schedule` to test lookup-only behaviour.
