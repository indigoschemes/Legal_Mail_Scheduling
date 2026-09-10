# Scheduled Mass Mail (GitHub Actions)

This is the GitHub-hosted twin of the local scheduler in `LG/`. Instead of
relying on this PC being on and Windows Task Scheduler running
`Run_Scheduler.bat`, GitHub's own servers run the exact same script on a
daily timer — see `.github/workflows/schedule.yml`.

The scheduling rules are identical to the local version: every row in
`Schedule.xlsx` has its own `Date` (a recurring `N Every month` spec, or
the exact deadline) and mail for that row goes out 5 days before it. The
`Date` column stays the real deadline — that's what shows up in the mail
via the `{date}` placeholder — while the script writes the actual send
date into a `Send Date` column each run, purely so you can see it in
Excel. See the main `README.md` in `LG/` for the full explanation of
`Date`, `Subject`/`Message`, recurring rows, and `Status` values — none
of that changes here.

`Subject`/`Message` can use two placeholders, filled in automatically when
the mail is built — write them as plain text, not an Excel formula (formulas
aren't evaluated by the headless script that runs on GitHub):
- `{date}` — the resolved send date for that row (`dd/mm/yyyy`), correctly
  handling recurring `N Every month` rows.
- `{label}` — the row's `Body Label` column value.

## One-time setup on GitHub

1. **Create a private repository** and push this folder's contents to it.
   (`git init`, `git add .`, `git commit`, add the remote, `git push`.)
2. **Add a repository secret** — repo → Settings → Secrets and variables →
   Actions → New repository secret:
   - Name: `LEGAL_MAILS_CREDENTIALS`
   - Value: a JSON object with the Gmail address to send from and its
     16-character Gmail App Password (see the main README's "Setting the
     sender" section for how to generate one), e.g.:
     ```json
     {"from_email": "legal@indigopaints.com", "app_password": "abcdefghijklmnop"}
     ```

   This is read by `mass_mail_scheduler.py` via an environment variable —
   nothing else to configure. `credentials.json` (used for local runs) is
   never used here and is excluded by `.gitignore`.
3. That's it. The workflow runs automatically every day at the time set
   in `schedule.yml` (default: 9:00 AM IST / 03:30 UTC — edit the `cron`
   line there to change it). You can also trigger a run manually any time
   from the repo's **Actions** tab → "Scheduled Mass Mail" → **Run workflow**.

## Keeping this folder in sync

Whenever you add rows, change content, or add attachments in your local
`Schedule.xlsx` / `Attachments`, copy the updated files into this folder
and push again (`git add`, `git commit`, `git push`) so the version
running on GitHub matches what you intend to send.

## Things to know

- **Timezone**: the workflow's cron schedule runs in UTC; the `cron` line
  in `schedule.yml` already accounts for the IST offset, but re-check it
  if you change the time.
- **Not to-the-minute precision**: GitHub can delay a scheduled run by a
  few minutes (occasionally more) under load — it still runs that day.
- **60-day auto-disable**: GitHub automatically disables a scheduled
  workflow after 60 days with no repository activity. A separate monthly
  `keepalive` job in `schedule.yml` guards against this by always making
  a small commit (a timestamp in `.keepalive`) on the 1st of every month,
  regardless of whether any mail was due — so the daily schedule never
  goes stale even during a long quiet stretch.
- **Attachments**: any file referenced in `Schedule.xlsx`'s `Attachment` /
  `Attachment2` columns must exist in this repo's `Attachments/` folder —
  GitHub's servers can't reach files on your PC.
