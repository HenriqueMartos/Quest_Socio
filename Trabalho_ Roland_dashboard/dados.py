import pandas as pd
from dash import Dash, dcc, html, Input, Output, no_update
import plotly.express as px  # biblioteca de gráficos
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import base64
from io import BytesIO
import nltk
from nltk.corpus import stopwords

# Baixa as stopwords em português (só precisa fazer uma vez)
nltk.download('stopwords')

# Função para converter letras de colunas (ex: "A", "B", "AA", "AB") em índices numéricos
def coluna_para_indice(coluna):
    indice = 0
    for letra in coluna:
        indice = indice * 26 + (ord(letra.upper()) - ord('A') + 1)
    return indice - 1  # Subtrai 1 porque o pandas usa indexação baseada em 0

# Carrega os dados do Excel
try:
    df = pd.read_excel(
        r"C:\Users\hfdma\Documents\ProgVscode\ProjetocustosLES\Trabalho_ Roland_dashboard\Question_Socio.xlsx",
        sheet_name="Sheet1"
    )
except FileNotFoundError:
    print("Arquivo Excel não encontrado. Verifique o caminho e o nome do arquivo.")
    df = pd.DataFrame()

if not df.empty:
    # Lista de colunas a serem removidas (baseado nas letras)
    colunas_para_remover = [
        "A", "B", "C", "D","E", "F", "G", "H", "J", "K", "M", "N", "O", "P", "Q", "S", "T", "V", "W", "Y", "Z",
        "AB", "AC", "AE", "AF", "AH", "AI", "AK", "AL", "AN", "AO", "AQ", "AR", "AT", "AU", "AW", "AX", "AZ", "BA",
        "BB", "BC", "BE", "BF", "BH", "BI", "BK", "BL", "BN", "BO", "BQ", "BR", "BT", "BU", "BW", "BX", "BZ", "CA",
        "CC", "CD", "CF", "CG", "CH", "CI", "CK", "CL", "CN", "CO", "CQ", "CR", "CT", "CU", "CW", "CX", "CZ", "DA",
        "DC", "DD", "DF", "DG", "DI", "DJ", "DL", "DM", "DO", "DP", "DR", "DS", "DU", "DV", "DW", "DX", "DZ", "EA",
        "EC", "ED", "EF", "EG", "EI", "EJ", "EK", "EL", "EN", "EO", "EQ", "ER", "ET", "EU", "EW", "EX", "EZ", "FA",
        "FC", "FD", "FE", "FF", "FH", "FI", "FK", "FL", "FN", "FO", "FQ", "FR", "FS", "FT", "FV", "FW", "FY", "FZ",
        "GB", "GC", "GE", "GF", "GH", "GI", "GK", "GL", "GM", "GN", "GP", "GQ", "GS", "GT", "GV", "GW", "GY", "GZ",
        "HA", "HB", "HD", "HE", "HG", "HH", "HJ", "HK", "HM", "HN", "HP", "HQ", "HS", "HT", "HV", "HW", "HX", "HY",
        "IA", "IB", "ID", "IE", "IG", "IH", "IJ", "IK", "IM", "IN", "IP", "IQ", "IR", "IS", "IU", "IV", "IX", "IY",
        "JA", "JB", "JC", "JD", "JF", "JG", "JI", "JJ", "JL", "JM", "JO", "JP", "JR", "JS", "JU", "JV", "JX", "JY",
        "KA", "KB", "KD", "KE", "KG", "KH", "KJ", "KK", "KM", "KN", "KP", "KQ", "KS", "KT", "KV", "KW", "KY", "KZ",
        "LB", "LC", "LE", "LF", "LH", "LI", "LK", "LL"
    ]

    # Converte as letras das colunas em índices e filtra os válidos
    indices_para_remover = [coluna_para_indice(col) for col in colunas_para_remover if coluna_para_indice(col) < len(df.columns)]

    # Remove as colunas com base nos índices
    df = df.drop(df.columns[indices_para_remover], axis=1)

    # Remove linhas completamente vazias
    df = df.dropna(how="all")

# Função para gerar a nuvem de palavras
def generate_wordcloud():
    if not df.empty and "Escreva algumas linhas sobre sua história e seus sonhos de vida" in df.columns:
        text = " ".join(df["Escreva algumas linhas sobre sua história e seus sonhos de vida"].dropna())
        stopwords_pt = set(stopwords.words('portuguese'))
        custom_stopwords = {"meu", "ter", "busco", "quero", "minha", "que", "um", "uma", "também", "para", "onde", "em", "área", "vida", "anos", "sonho", "outros", "fazer", "possa","sempre","duas", "ano", "nisso"}
        stopwords_pt.update(custom_stopwords)
        words = text.split()
        filtered_words = [word for word in words if word.lower() not in stopwords_pt]
        filtered_text = " ".join(filtered_words)

        wordcloud = WordCloud(width=800, height=400, background_color="white", max_words=100, colormap="viridis", stopwords=stopwords_pt).generate(filtered_text)

        buffer = BytesIO()
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation="bilinear")
        plt.axis("off")
        plt.savefig(buffer, format="png")
        plt.close()

        return base64.b64encode(buffer.getvalue()).decode("utf-8")
    return None

# Inicializa o app Dash
app = Dash(__name__, assets_folder="assets")

app.layout = html.Div([
    html.H1("Dashboard Socioeconômico - FATEC Franca", className="dashboard-title"),
    html.P("Escolha uma das opções abaixo.", className="dashboard-description"),
    html.Div(
        dcc.Dropdown(
            id="dropdown-column",
            options=[{"label": col, "value": col} for col in df.columns] if not df.empty else [],
            value=df.columns[0] if not df.empty else None,
            clearable=False,
            className="dropdown"
        ),
        className="dropdown-container"
    ),
    dcc.Graph(id="graph"),
    html.Div(id="wordcloud-container", className="wordcloud-container")
])


@app.callback(
    Output("graph", "figure"),
    Input("dropdown-column", "value")
)
def update_graph(selected_column):
    if selected_column == "Escreva algumas linhas sobre sua história e seus sonhos de vida":
        # Retorna uma figura vazia com mensagem informativa
        fig = px.pie(title="Este campo exibe apenas a nuvem de palavras.", template="plotly_dark")
        fig.update_layout(showlegend=False)
        return fig

    if not df.empty and selected_column in df.columns:
        value_counts = df[selected_column].value_counts().reset_index()
        value_counts.columns = ['Categoria', 'Contagem']

        fig = px.pie(
            value_counts,
            names='Categoria',
            values='Contagem',
            title=f"Distribuição de {selected_column}",
            hole=0.4,  # transforma em gráfico de rosca (opcional)
            color_discrete_sequence=["#A259FF", "#6C63FF", "#7B61FF", "#9D4EDD", "#B983FF", "#D0A2F7"]

        )

        fig.update_traces(
            textinfo='percent+label',  # mostra porcentagem e nome da categoria
            textfont_size=14,
            pull=[0.03]*len(value_counts)  # leve destaque visual
        )

        fig.update_layout(
            template='plotly_dark',
            showlegend=True,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
        )

        return fig

    return px.pie(title="Sem dados disponíveis", template="plotly_dark")



@app.callback(
    Output("wordcloud-container", "children"),
    Input("dropdown-column", "value")
)
def update_wordcloud(selected_column):
    if selected_column == "Escreva algumas linhas sobre sua história e seus sonhos de vida":
        image_base64 = generate_wordcloud()
        if image_base64:
            return [
                html.H2("Nuvem de Palavras - Sonhos e Histórias"),
                html.Img(src=f"data:image/png;base64,{image_base64}")
            ]
        else:
            return html.P("Nenhum dado disponível para gerar a nuvem de palavras.")
    return None

if __name__ == "__main__":
    app.run(debug=False)
