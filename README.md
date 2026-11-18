# Wedding Expo 2027 — Spin & Win

## Start the game

Requires Python 3.10+ and the `artifact_tool` Python package for `.xlsx` export. The web server and SQLite storage use Python's standard library. `artifact_tool` is needed only when exporting Excel.

Mac/Linux:
```bash
export ADMIN_PASSWORD='choose-a-long-private-password'
export EXPO_DATE='2027-02-01' # change to your actual exhibition date
python server.py
```

Windows PowerShell:
```powershell
$env:ADMIN_PASSWORD='choose-a-long-private-password'
$env:EXPO_DATE='2027-02-01'
python server.py
```

Open `http://localhost:8080` on the exhibition computer. To play from phones, deploy to a HTTPS hosting service with a **persistent disk** for the SQLite database and a public HTTPS domain. Do not share a `localhost` QR code: it only works on the computer hosting the game. Once deployed, generate a QR code pointing to your public HTTPS URL, e.g. with a QR generator or Python `qrcode` package.

## Admin access

- View records: `https://YOUR-DOMAIN/admin/entries`
- Download live Excel export: `https://YOUR-DOMAIN/admin/export.xlsx`
- Update a claim, booking status, staff follow verification and notes: POST JSON to `/admin/update` (Basic authentication) with `id`, `claim_status`, `booking_status`, `follow_verified`, `notes`. A simple administrative dashboard can be added later.

Both admin routes use browser HTTP Basic Authentication: username can be anything; password must equal `ADMIN_PASSWORD`. **HTTPS is mandatory** for public deployment because Basic Authentication transmits reusable credentials.

## Reliability and fair allocation

- SQLite is the authoritative record, stored in `wedding_game.sqlite3`; Excel is a downloadable snapshot of committed records.
- A single atomic transaction checks duplicates, allocates a remaining prize, decrements inventory, and saves the record.
- Concurrent visitors cannot claim the same final prize. A failed transaction will not consume a prize.
- Email+event date and phone+event date are unique; staff should resolve legitimate shared-contact cases.
- Default prize inventory: 1 free photoshoot, 3 x 20%, 6 x 15%, 20 x 10%, 2 event trailers, 3 DJI behind-the-scenes, 5 extra 30-minute coverage = 40 total. Edit `PRIZES` before the first launch. Existing inventory is intentionally not overwritten on restart.
- Allocation is random weighted by the **remaining** prize counts; every eligible entry wins until inventory is exhausted. Publish the rules and allocation method before the exhibition.
- The social follow is self-attested online. Staff must check the account in person; there is no automatic social-platform verification.
- The frontend submits and saves before the wheel animates. If the network fails after saving but before response, a repeat submission will be flagged as duplicate; staff can look up the entry rather than allocating another prize.

## Security, privacy and operating checklist

- Replace the sample terms with your final full terms, organiser identity, actual social handles, and privacy notice.
- Ensure a lawful basis for processing registration data and a separate marketing opt-in; set a retention period and handle deletion requests.
- Use a reputable HTTPS host with secure server-side storage, restricted admin credentials, regular encrypted backups and a documented restore test.
- Avoid running the public server directly on an unsecured venue Wi-Fi network. Use HTTPS hosting and test it on both tablet and phone.
- Export a copy of the register periodically during the event. Excel exports are **not** live writable databases.
- Verify prize-specific deliverables and paid-booking restrictions, especially the free photoshoot, before launch.
- The 14-day claim and 30-day booking deadlines are calculated from `EXPO_DATE` in the Excel export. Set this correctly.
- Before a live event, add rate limiting, anti-bot controls, monitoring and a recovery workflow if traffic or abuse is expected. This starter is not a fully hardened production service.
