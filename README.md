# Health Analytics Case Generator (HINF1009)

Generates a client request, a synthetic healthcare dataset with built-in relationships,
and an instructor answer key with beginner Excel steps.

## How it works
1. Claude writes the client memo, tasks, and a JSON data design (variables + relationships).
2. Python generates the rows from that design (NumPy), optionally adding messy data.
3. Python computes the true results; Claude turns them into an answer key with Excel formulas.
4. Downloads: a student workbook (memo + data) and an instructor workbook
   (answer key, results tables, hidden relationships, clean data).

## Run locally
    pip install -r requirements.txt
    cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # add your API key
    streamlit run app.py

## Deploy free on Streamlit Community Cloud
1. Put this folder in a GitHub repository (secrets.toml is excluded by .gitignore).
2. At share.streamlit.io, create an app pointing to app.py.
3. In App settings > Secrets, paste ANTHROPIC_API_KEY (and optionally APP_PASSWORD).

## Adding scenarios
Edit the SCENARIOS dictionary at the top of app.py, or use "Custom scenario..." in the app.
The seed shown after each run recreates the same dataset.
