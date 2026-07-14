# HiFleet Weather Dashboard

Automatic daily weather dashboard for 76 ships, generated from HiFleet weather email.

## Features
- 76 ships (excluding ORE CHINA, ORE DONGJIAKOU, ORE HEBEI, ORE SHANDONG)
- 7-day wind / visibility / wave forecast
- Real-time alerts: wind > 6 m/s OR wave >= 3m
- Multi-dimensional wave ≥3m analysis

## Auto-run
- GitHub Actions runs daily at 23:50 CST
- Fetches latest HiFleet email via QQ Mail IMAP
- Generates HTML dashboard
- Publishes to GitHub Pages

## Secrets Required
- `HIFLEET_EMAIL_USER`: QQ email (xuehaifeng666@qq.com)
- `HIFLEET_EMAIL_AUTH`: IMAP authorization code
- `QQ_APP_PWD`: QQ IMAP app password (alternative)

## Live Dashboard
https://xuehaifeng80-alt.github.io/hifleet-weather/
