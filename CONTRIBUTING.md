# Contributing

Thank you for helping. This project is maintained by Oriental Horizon Hospitality.

## Ways to help

- **A fix** when a platform changes its page (usually the extraction code in `scripts/<platform>.js`)
- **Other platforms**, other languages for the keyword lists, translations of the report
- **Report and chart improvements**

## How

1. Fork the repository, make a branch, make your change, open a pull request.
2. Say what the change does and, for scrapers, **what you tested it on and when**.
3. Run `python tests/test_core.py` and make sure it prints the all-passed line.
4. Keep pull requests small and focused.

## Rules

1. **Stop on a login page, CAPTCHA or bot check.** Never add code that bypasses it or retries aggressively.
2. **Never ask for or store passwords.** The user logs in; keep session files out of git.
3. Be slow and gentle: pauses between pages, small volumes.
4. Do not claim a platform works unless you ran it. Update the status table in `README.md` only with what you tested.
5. Never commit real guest data you are not allowed to share. Sample data must be invented.
6. Do not add code for sites whose terms forbid automated access without permission.
