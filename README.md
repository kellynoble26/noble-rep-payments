# Noble REP Payment Sheet 2026

Passcode-locked GitHub Pages site for Noble Windows and Doors REP payments.

- Live: https://kellynoble26.github.io/noble-rep-payments/
- Source of truth: OneDrive Documents / `Noble_WD_REP_Payment_Sheet_2026.xlsx`
- Only the encrypted `index.html` is published. The plaintext workbook and HTML stay off this branch tip used for Pages.

Refresh from `/workspace/rep-web/refresh.sh` (requires `FIELD_LEDGER_PASSCODE`; never commit the passcode).

Monies received (GC and WD together) is `received.html`, linked from the passcode screen. Rebuild it with `tools/build_monies_received.py`. The same `FIELD_LEDGER_PASSCODE` seals the page; the script never prints or stores the passcode. Pages uploads `received.html` only after that seal. An unlocked file on a review branch is omitted from the Pages artifact.
