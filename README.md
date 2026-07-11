# Options Brokerage Studio

A Streamlit dashboard for analysing MT Quant option backtest exports. It calculates premium turnover, Zerodha brokerage, statutory charges, net P&L, India VIX bands, underlying-change filters, and downloadable Excel reports.

## Run locally

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push `app.py`, `requirements.txt`, `.streamlit/config.toml`, `.gitignore`, and this README to a GitHub repository.
2. Sign in to [Streamlit Community Cloud](https://share.streamlit.io/) with GitHub.
3. Select **Create app** and choose the repository and branch.
4. Set the entrypoint to `app.py` and deploy.

Uploaded trading files are processed in memory and are not included in this repository.
