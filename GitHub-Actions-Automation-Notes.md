# GitHub Actions Automation — Notes from This Project

Reusable notes on the patterns used to turn `mass_mail_scheduler.py` into a
mail system that runs entirely on GitHub's servers, with no PC, Excel, or
person needing to be present. Kept generic enough to reapply to any future
"run this script on a schedule / send data somewhere automatically" project.

---

## 1. The core idea: a workflow file

Everything lives in one YAML file: `.github/workflows/schedule.yml`. GitHub
watches this path in every repo; any file there is a "workflow" it can run
on its own servers (Ubuntu/Windows/macOS runners), for free on public repos
and with a monthly free quota on private ones.

A workflow has:
- **Triggers** (`on:`) — what causes it to run.
- **Jobs** (`jobs:`) — one or more independent units of work, each running
  on a fresh, disposable virtual machine.
- **Steps** — the ordered commands inside a job (checkout code, install
  deps, run a script, commit results).

## 2. Triggers used here

```yaml
on:
  schedule:
    - cron: "30 3 * * *"   # daily
    - cron: "0 0 1 * *"    # monthly
  workflow_dispatch: {}
```

- **`schedule` / cron**: GitHub's cron is always **UTC**, regardless of
  where you are. Convert your local time yourself (IST = UTC + 5:30):
  | Local (IST) | UTC   | Cron          |
  |-------------|-------|---------------|
  | 9:00 AM     | 03:30 | `30 3 * * *`  |
  | 12:30 PM    | 07:00 | `0 7 * * *`   |
  - A workflow can list **multiple `cron:` entries** under one `schedule:`
    key — each fires independently, and each triggered run carries
    `github.event.schedule` set to the exact cron string that fired it, so
    a job can check `if: github.event.schedule == '...'` to react only to
    one of them (used below for the keep-alive job).
  - Scheduled runs are **not precise to the minute** — GitHub can delay a
    few minutes under load, occasionally longer. Fine for daily/monthly
    jobs, not for anything time-critical.
- **`workflow_dispatch`**: adds a manual **"Run workflow"** button on the
  repo's Actions tab — essential for testing without waiting for the cron.

## 3. Secrets — never commit credentials

Repo → **Settings → Secrets and variables → Actions → Repository secrets**.
Referenced in the workflow as `${{ secrets.NAME }}`, injected only as an
environment variable at run time — never visible in logs, never stored in
the repo.

```yaml
env:
  LEGAL_MAILS_CREDENTIALS: ${{ secrets.LEGAL_MAILS_CREDENTIALS }}
```

We packed the sender email + app password into **one JSON secret** rather
than two separate ones, and had the Python side do
`json.loads(os.environ["LEGAL_MAILS_CREDENTIALS"])` — simpler to manage
than keeping multiple secrets in sync, and easy to extend later.

The corresponding local-dev fallback (`credentials.json`) is listed in
`.gitignore` so a local secrets file can never be committed by accident.

## 4. Permissions

```yaml
permissions:
  contents: write
```

Needed because the job pushes a commit back to the repo (see §5). Default
token permissions are read-only; this opts the job's auto-generated
`GITHUB_TOKEN` into write access, scoped to just this repo, no PAT needed.

## 5. The "commit results back" pattern

The script marks each row `Sent (...)` in `Schedule.xlsx` after sending,
then the workflow commits that change back to the repo:

```yaml
- name: Commit updated Schedule.xlsx
  run: |
    if ! git diff --quiet -- Schedule.xlsx; then
      git config user.name "github-actions[bot]"
      git config user.email "github-actions[bot]@users.noreply.github.com"
      git add Schedule.xlsx
      git commit -m "Update Schedule.xlsx after scheduled run"
      git push
    else
      echo "No changes -- nothing to commit."
    fi
```

This is the idempotency mechanism: without it, the same row would be sent
again on every run since nothing records that it already went out. The
`git diff --quiet` guard avoids empty commits when nothing changed.
`github-actions[bot]` is GitHub's conventional identity for bot commits.

**Caution**: if two runs overlap (e.g. a manual trigger while the daily
one is still running), both may try to push and one will be rejected —
not an issue for a once-a-day job, but worth knowing for higher-frequency
schedules.

## 6. Headless execution has no GUI apps — plan around it

The runner has no Excel, no display, nothing interactive. This bit us
directly: an **Excel formula** in a cell (e.g.
`="<p>Dear Team,</p>"&...&TEXT($L2,"dd/mm/yyyy")&...`) is never evaluated
by Python's `openpyxl` in headless mode — reading the cell returns the
**literal formula text**, not a computed value. Silent, no error, just
wrong output (the raw formula string would have gone out as the mail body).

**Takeaway for any future project**: if a spreadsheet/file mixes formulas
and data that a headless script will read, either (a) have the script do
that computation itself in code, or (b) always read with
`data_only=True` *and* guarantee the file was last saved by an app that
actually calculated it (fragile — breaks the moment anything else
resaves the file without recalculating, e.g. a script round-tripping it).
We chose (a): moved the "date 5 days before deadline" and "insert this
label into the message" logic into Python, and reduced the spreadsheet
to plain-text placeholders — see §7.

## 7. Placeholder templating instead of formulas

`Subject`/`Message` cells contain plain text with tokens:

```
<p>Dear Team,</p>
<p>{label} due date - {date} </p>
```

The script fills these in right before sending:

```python
for placeholder, value in (("{date}", row.get("send_date_str", "")),
                            ("{label}", row.get(BODY_LABEL_COL, ""))):
    subject = subject.replace(placeholder, value)
    body = body.replace(placeholder, value)
```

Generalizes to any "mail merge"-style need — keep the data model as plain
text + placeholders, keep all computation in code, never in the file the
code also reads.

## 8. Deadline vs. send-date separation

Two distinct dates were easy to conflate but needed to stay separate:
- **`Date`** column — the actual deadline, shown to the recipient via
  `{date}`. Never modified by the send-timing logic.
- **`Send Date`** column — auto-written by the script every run
  (`Date` − `SEND_DAYS_BEFORE`), purely informational, so a human can see
  in the spreadsheet exactly when each row will fire — without that
  logic being hidden inside a formula (see §6) or requiring manual entry.

A recurring spec like `"15 Every month"` is resolved fresh on every run
(clamped to the actual last day of shorter months) rather than resolved
once and gone stale — recomputing cheap logic beats caching it in a cell.

## 9. Preventing the 60-day auto-disable

GitHub automatically disables a repo's scheduled workflows after **60
days with no repository activity**. A commit counts as activity; a
workflow merely *running* without committing anything does not (the
"nothing due today" days from §5 don't reset the clock). This matters for
any low-frequency schedule where long gaps between real events are normal.

**Fix**: a second job on its own monthly cron that unconditionally commits
a timestamp file:

```yaml
keepalive:
  if: github.event.schedule == '0 0 1 * *'
  runs-on: ubuntu-latest
  permissions:
    contents: write
  steps:
    - uses: actions/checkout@v4
    - name: Record keep-alive commit
      run: |
        git config user.name "github-actions[bot]"
        git config user.email "github-actions[bot]@users.noreply.github.com"
        date -u +"%Y-%m-%dT%H:%M:%SZ" > .keepalive
        git add .keepalive
        git commit -m "Monthly keep-alive commit"
        git push
```

And the main job is scoped to *not* run on that cron, so the two don't
double-fire:

```yaml
send:
  if: github.event_name != 'schedule' || github.event.schedule == '30 3 * * *'
```

Reusable rule of thumb: **any workflow that might go 60+ days without a
"real" commit needs an artificial one.**

## 10. General checklist for a future "automate this on GitHub" project

1. Does the script need secrets? → repository secrets, injected as env
   vars, never written to any tracked file.
2. Does the script need to remember state between runs (e.g. "already
   sent")? → write that state back into the repo and commit it
   (`contents: write` permission required).
3. Does the input data mix formulas/computed values with raw data? → move
   the computation into the script itself; keep the file just data +
   placeholders.
4. Is the schedule infrequent or bursty (could go 60+ days quiet)? → add
   a keep-alive job on its own cron.
5. Want to test without waiting for the schedule? → add
   `workflow_dispatch: {}` for a manual "Run workflow" button.
6. Remember cron is UTC — convert your local trigger time by hand.
