# Deploying the Morning Financial Wire to Streamlit Cloud

Streamlit Cloud cannot reach Outlook, so the app has two halves:

- **Your PC** reads Outlook and pushes the results to Snowflake (`push_wire.py`).
- **Streamlit Cloud** reads Snowflake and fetches the public news feeds live (`streamlit_app.py`).

## 1. Snowflake (once)
1. Make a key pair (PowerShell, in `~\.snowflake`):
   `openssl genrsa 2048 | openssl pkcs8 -topk8 -nocrypt -inform PEM -out fin_wire_svc.p8`
   `openssl rsa -in fin_wire_svc.p8 -pubout -out fin_wire_svc.pub`
2. Paste the public key body into `snowflake_setup.sql` and run it in Snowsight as ACCOUNTADMIN.

## 2. Your PC
Create `financial_wire\.env` (gitignored):

    FINWIRE_SF_ACCOUNT=JSA-ANALYTICS
    FINWIRE_SF_USER=FIN_WIRE_SVC
    FINWIRE_SF_PRIVATE_KEY_PATH=C:\Users\KoltenPostin\.snowflake\fin_wire_svc.p8

Test, then schedule weekdays ~6:00 AM CT in Task Scheduler (Outlook must be running):

    python push_wire.py --dry
    python push_wire.py

## 3. GitHub + Streamlit Cloud
1. Push `financial_wire` to a **private** repo under JSA-Dashboards. The `.gitignore` keeps the fetched email text out of the repo.
2. share.streamlit.io -> New app -> that repo -> `streamlit_app.py`.
3. App settings -> Secrets:

       [snowflake]
       account = "JSA-ANALYTICS"
       user = "FIN_WIRE_SVC"
       role = "FIN_WIRE_ROLE"
       warehouse = "COMPUTE_WH"
       database = "FIN_WIRE"
       schema = "PUBLIC"
       private_key = """-----BEGIN PRIVATE KEY-----
       ...
       -----END PRIVATE KEY-----"""

4. **Share -> restrict viewing to your own email.** PJ/StoneX, Bloomberg and WSJ text is licensed for your own reading, so the app must not be public.
