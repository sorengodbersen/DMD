import dash
import dash_bootstrap_components as dbc
from dash import dcc, html, dash_table, Input, Output
import pandas as pd


# Load the Excel file
excel_file_path = 'DMD.xlsx'
df = pd.read_excel(excel_file_path)

# Initialize Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = 'DMD Dashboard'

# App layout
app.layout = dbc.Container([
    html.H1("DMD Dashboard", className="text-center my-4"),

    dbc.Row([
        dbc.Col([
            dash_table.DataTable(
                id='data-table',
                columns=[{"name": i, "id": i} for i in df.columns],
                data=df.to_dict('records'),
                filter_action='native',
                sort_action='native',
                page_action='native',
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left', 'padding': '5px'},
                style_header={'backgroundColor': '#f8f9fa', 'fontWeight': 'bold'},
            )
        ], width=12)
    ]),

    dbc.Row([
        dbc.Col([
            html.Button("Download Data", id="btn-download", className="btn btn-primary mt-3"),
            dcc.Download(id="download-data")
        ])
    ])
], fluid=True)

# Callback for downloading data
@app.callback(
    Output("download-data", "data"),
    Input("btn-download", "n_clicks"),
    prevent_initial_call=True
)
def download_data(n_clicks):
    return dcc.send_file(excel_file_path)

# Run the app
if __name__ == '__main__':
    app.run(host="0.0.0.0", port=8080)