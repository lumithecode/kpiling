# KPIling KPI Scorecard

A Streamlit dashboard for weekly operating KPIs, data quality checks, target management, and formatted Excel reporting. The app creates a deterministic 12-week sample dataset with one missing CSAT value and one support-ticket outlier so the health checks are visible on first run.

## Run locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
streamlit run app.py
```

The app opens in your browser at the local Streamlit URL. In VS Code, select the environment where the requirements were installed if the editor prompts for an interpreter.

## What is included

- **Dashboard:** six KPIs, weekly/monthly views, target status, period comparison, and a selectable 12-week trend.
- **Raw Data:** editable weekly records, CSV upload, and automatic calculation of week number and conversion rate.
- **Data Health:** missing-value, duplicate-date, outlier, and freshness checks with a health score and action log.
- **Targets & Definitions:** editable weekly/monthly targets and metric directionality saved per user.
- **Documentation:** metric aggregation rules and guidance for replacing sample data with a production source.
- **Excel export:** a timestamped five-sheet workbook with a cover summary, scorecard, raw history, health checks, and definitions.

## Deployment

The app is pinned to Python 3.14.7 in `.python-version` and exact dependency versions in `requirements.txt`. Choose a hosting runtime that supports this Python version. Run the test suite before deployment:

```powershell
python -m pytest
```

Local development uses SQLite at `data/kpi_scorecard.db`. That folder is git-ignored. For a hosted deployment, use a managed PostgreSQL database and provide its URL as `DATABASE_URL`, for example with the `postgresql+psycopg://` scheme. Environment variables work on all hosts; Streamlit Community Cloud users can put the same keys in the app's Secrets settings. Do not use a local SQLite file as persistent cloud storage on an ephemeral host.

Configure OIDC using your identity provider and the safe example at `.streamlit/secrets.toml.example`. For Streamlit Community Cloud, add the example's root keys and `[auth]` section to **App settings → Secrets**. For other hosts, use environment variables for `DATABASE_URL`, `REQUIRE_AUTH`, and `ALLOWED_EMAILS`, and the host's secret manager for OIDC settings. Set `redirect_uri` to the deployed URL ending in `/oauth2callback`, and replace every placeholder. Never commit `secrets.toml` or real credentials.

Set `REQUIRE_AUTH=true` in the deployment environment to fail closed unless OIDC is configured. Optionally set `ALLOWED_EMAILS` to a comma-separated email allowlist. Signed-in users have isolated workspaces keyed by a hash of their OIDC issuer and subject. OIDC authenticates users; the allowlist and per-user workspace scope provide this app's access control. CSV uploads are limited to 20 MB.

When changing the hosted URL, update both the identity provider's allowed callback URL and `redirect_uri`. Back up the managed database and use TLS-protected connection settings provided by your database host.

## Sample data and targets

The sample data is generated in `generate_sample_data()` in `app.py`. It ends with the latest complete Monday-to-Sunday week. In the **Raw Data** tab, upload a CSV with one row per week. Required headers are `Week Start`, `Revenue`, `Active Users`, `Signups`, `Conversions`, `Support Tickets Open`, `CSAT`, and `Headcount`. Use a date such as `2026-09-28` for `Week Start`; `Week End`, `Week #`, `Conversion Rate`, and `Notes` are optional. The app derives `Week End`, `Week #`, and `Conversion Rate` when missing and recalculates week number and conversion rate after edits. You can also edit cells or add/remove rows directly in the table. Use **Restore sample data** in the sidebar to return to the original demo.

KPI status rules are implemented in `status_for_value()`; aggregation and comparison logic is in `calculate_kpis()`.
