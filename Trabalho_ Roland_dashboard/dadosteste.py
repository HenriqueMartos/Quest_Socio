import pandas as pd
from dash import Dash, dcc, html, Input, Output
import plotly.express as px
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import base64
from io import BytesIO
import nltk
from nltk.corpus import stopwords

nltk.download('stopwords')

# === CONFIGURAÇÃO INICIAL ===

# Caminho do arquivo Excel
file_path = r"C:\Users\hfdma\Documents\ProgVscode\ProjetocustosLES\Trabalho_ Roland_dashboard\Question_Socio.xlsx"

try:
    df = pd.read_excel(file_path, sheet_name="Sheet1")
except FileNotFoundError:
    print("Arquivo Excel não encontrado.")
    df = pd.DataFrame()

# === REMOVER COLUNAS SENSÍVEIS SOMENTE NA LEITURA ===

colunas_sensiveis = ["Timestamp", "Nome completo", "Telefone", "E-mail"]
colunas_visiveis = [col for col in df.columns if col not in colunas_sensiveis]
df_processado = df[colunas_visiveis].copy()

# === FUNÇÃO PARA GERAÇÃO DE NUVEM DE PALAVRAS ===

def generate_wordcloud(df_base):
    coluna_alvo = "Escreva algumas linhas sobre sua história e seus sonhos de vida"
    if not df_base.empty and coluna_alvo in df_base.columns:
        text = " ".join(df_base[coluna_alvo].dropna())
        stopwords_pt = set(stopwords.words("portuguese"))
        custom = {"meu", "ter", "busco", "quero", "minha", "que", "um", "uma", "também", "para",
                  "onde", "em", "área", "vida", "anos", "sonho", "outros", "fazer", "possa", "sempre", "duas", "ano", "nisso"}
        stopwords_pt.update(custom)
        words = text.split()
        filtered_words = [word for word in words if word.lower() not in stopwords_pt]
        filtered_text = " ".join(filtered_words)

        wordcloud = WordCloud(width=800, height=400, background_color="black", colormap="plasma").generate(filtered_text)
        buffer = BytesIO()
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation="bilinear")
        plt.axis("off")
        plt.tight_layout()
        plt.savefig(buffer, format="png")
        plt.close()
        return base64.b64encode(buffer.getvalue()).decode("utf-8")
    return None

# === DASH APP ===

app = Dash(__name__, assets_folder="assets")  # Certifique-se de ter assets/style.css para o dark mode

app.layout = html.Div([
    html.H1("Dashboard Socioeconômico - FATEC Franca", className="dashboard-title"),
    html.P("Escolha uma das opções abaixo.", className="dashboard-description"),
    
    html.Div(
        dcc.Dropdown(
            id="dropdown-column",
            options=[{"label": col, "value": col} for col in df_processado.columns],
            value=df_processado.columns[0] if not df_processado.empty else None,
            clearable=False,
            className="dropdown"
        ),
        className="dropdown-container"
    ),
    
    dcc.Graph(id="graph"),
    html.Div(id="wordcloud-container", className="wordcloud-container")
])

# === CALLBACKS ===

@app.callback(
    Output("graph", "figure"),
    Input("dropdown-column", "value")
)
def update_graph(selected_column):
    if not df_processado.empty and selected_column in df_processado.columns:
        # Gráfico de pizza
        df_temp = df_processado[selected_column].dropna()
        fig = px.pie(df_temp.value_counts().reset_index(), names='index', values=selected_column,
                     title=f"Distribuição de {selected_column}")
        fig.update_layout(template="plotly_dark")
        return fig
    return px.pie(title="Sem dados disponíveis", template="plotly_dark")

@app.callback(
    Output("wordcloud-container", "children"),
    Input("dropdown-column", "value")
)
def update_wordcloud(selected_column):
    if selected_column == "Escreva algumas linhas sobre sua história e seus sonhos de vida":
        image_base64 = generate_wordcloud(df_processado)
        if image_base64:
            return [
                html.H2("Nuvem de Palavras - Sonhos e Histórias"),
                html.Img(src=f"data:image/png;base64,{image_base64}", className="wordcloud-img")
            ]
        else:
            return html.P("Nenhum dado disponível para gerar a nuvem de palavras.")
    return None

# === INICIALIZA O SERVIDOR ===

if __name__ == "__main__":
    app.run(debug=False)
