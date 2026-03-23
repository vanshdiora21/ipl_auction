# IPL 2026 Fantasy Points Tracker

Automatically fetches IPL scorecard data from CREX and logs fantasy points
for your player list into `players.xlsx` every day at **8:00 PM IST**.

## How it works

1. GitHub Actions runs the script daily at 8pm IST via cron
2. The script fetches today's IPL scorecard from crex.com
3. Fantasy points are calculated using the standard Dream11/CREX T20 system
4. Results are written into `players.xlsx` under today's date column
5. The updated Excel is automatically committed back to this repo

## Your player list

Edit `players.xlsx` directly:
- **Column A**: Player Name (e.g. `Abhishek Sharma`)
- **Column B**: Team (e.g. `SRH`)
- Date columns fill in automatically

## Running locally

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 fantasy_points_tracker.py
```

## Triggering manually

Go to **Actions** tab on GitHub → select **Fantasy Points Tracker** → click **Run workflow**.

## Fantasy point system (T20)

| Action | Points |
|--------|--------|
| Playing XI | +4 |
| Per run | +1 |
| Per boundary (4) | +1 |
| Per six | +2 |
| Half-century | +8 |
| Century | +16 |
| Duck (non-bowler) | -2 |
| SR ≥ 170 (10+ balls) | +6 |
| SR ≥ 150 | +4 |
| SR ≥ 130 | +2 |
| SR < 60 | -2 |
| SR < 50 | -4 |
| SR < 40 | -6 |
| Per wicket | +25 |
| LBW / Bowled bonus | +8 |
| 3-wicket haul | +4 |
| 4-wicket haul | +8 |
| 5-wicket haul | +16 |
| Maiden over | +12 |
| Economy < 4 (2+ overs) | +6 |
| Economy < 5 | +4 |
| Economy < 7 | +2 |
| Economy > 12 | -2 |
| Economy > 13 | -4 |
| Economy > 14 | -6 |
| Catch | +8 |
| Stumping | +12 |
| Direct run-out | +12 |
