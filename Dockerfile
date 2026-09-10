FROM python:3.12-slim

# Non-root user for IBM Cloud Code Engine (required)
RUN useradd -m -u 1001 appuser

WORKDIR /app

# Install dependencies first (layer cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source
COPY bench_forecast_dashboard.py .

# Copy data files (Excel source + override state)
COPY "NA Bench Forecast.xlsx" "./NA Bench Forecast.xlsx"
COPY bench_overrides.json .

# Copy static assets
COPY "IBM LOGO.png" "./IBM LOGO.png"

# Streamlit runtime config via env
ENV STREAMLIT_SERVER_PORT=8080
ENV STREAMLIT_SERVER_ADDRESS=0.0.0.0
ENV STREAMLIT_SERVER_HEADLESS=true
ENV STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# Streamlit secrets are injected at runtime via Code Engine secret mount:
#   /app/.streamlit/secrets.toml
# Keys used by this app:
#   [forecast] actuals_through = <int>
#   [auth]     password = <string>   (optional access gate)
#   [admin]    password = <string>   (optional audit log gate)
#   [lock]     editing  = <bool>     (optional read-only mode)

EXPOSE 8080

USER 1001

CMD ["streamlit", "run", "bench_forecast_dashboard.py", "--server.port=8080", "--server.address=0.0.0.0"]
