import pandas as pd
from dash import Dash, dcc, html, Input, Output, no_update
import plotly.express as px #biblioteca graficos
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import base64
from io import BytesIO
import nltk
from nltk.corpus import stopwords

#perguntar ao chat sobre a instalação das bibliotecas e integrar no vs
#strimlit biblioteca

# Baixa as stopwords em português (só precisa fazer uma vez)
nltk.download('stopwords')

# Função para converter letras de colunas (ex: "A", "B", "AA", "AB") em índices numéricos
def coluna_para_indice(coluna):
    indice = 0
    for letra in coluna:
        indice = indice * 26 + (ord(letra.upper()) - ord('A') + 1)
    return indice - 1  # Subtrai 1 porque o pandas usa indexação baseada em 0

# Carrega os dados do Excel
df = pd.read_excel(r"C:\Users\hfdma\Documents\ProgVscode\ProjetocustosLES\Trabalho_ Roland_dashboard\Question_Socio.xlsx", sheet_name="Sheet1")

# Lista de colunas a serem removidas (baseado nas letras)
colunas_para_remover = [
    "A", "B", "C", "D", "F", "G", "H", "J", "K", "M", "N", "O", "P", "Q", "S", "T", "V", "W", "Y", "Z",
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

# Converte as letras das colunas em índices
indices_para_remover = [coluna_para_indice(col) for col in colunas_para_remover]

# Remove as colunas com base nos índices
df = df.drop(df.columns[indices_para_remover], axis=1)

# Remove linhas completamente vazias
df = df.dropna(how="all")

# Função para gerar a nuvem de palavras
def generate_wordcloud():
    if "Escreva algumas linhas sobre sua história e seus sonhos de vida" in df.columns:
        # Junte todos os textos em uma única string
        text = " ".join(df["Escreva algumas linhas sobre sua história e seus sonhos de vida"].dropna())

        # Lista de stopwords em português
        stopwords_pt = set(stopwords.words('portuguese'))

        # Adicione palavras personalizadas que você quer remover
        custom_stopwords = {"meu", "ter", "busco", "quero", "minha", "que", "um", "uma", "também", "para", "onde", "em", "área", "vida", "anos", "sonho", "outros", "fazer", "possa","sempre","duas", "ano", "nisso"}
        stopwords_pt.update(custom_stopwords)

        # Filtra as stopwords do texto
        words = text.split()
        filtered_words = [word for word in words if word.lower() not in stopwords_pt]
        filtered_text = " ".join(filtered_words)

        # Cria a nuvem de palavras
        wordcloud = WordCloud(
            width=800,
            height=400,
            background_color="white",
            max_words=100,
            colormap="viridis",
            stopwords=stopwords_pt  # Remove as stopwords diretamente no WordCloud
        ).generate(filtered_text)

        # Salva a nuvem de palavras em uma imagem temporária
        buffer = BytesIO()
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation="bilinear")
        plt.axis("off")
        plt.savefig(buffer, format="png")
        plt.close()

        # Converte a imagem para base64
        return base64.b64encode(buffer.getvalue()).decode("utf-8")
    else:
        return None

# Inicializa o app Dash e referencia a pasta assets
app = Dash(__name__, assets_folder="assets")

# Layout do dashboard
app.layout = html.Div([
    # Título com estilo personalizado
    html.H1("Dashboard Socioeconômico - FATEC Franca", className="dashboard-title"),
    
    # Parágrafo com estilo personalizado
    html.P("Escolha uma das opções abaixo.", className="dashboard-description"),
    
    # Dropdown para selecionar a coluna a ser analisada
    html.Div(
        dcc.Dropdown(
            id="dropdown-column",
            options=[{"label": col, "value": col} for col in df.columns],
            value=df.columns[0],  # Valor inicial
            clearable=False,
            className="dropdown"  # Aplica a classe CSS
        ),
        className="dropdown-container"  # Aplica a classe CSS
    ),
    
    # Gráfico interativo
    dcc.Graph(id="graph"),

    # Nuvem de palavras (só aparece se a coluna específica for selecionada)
    html.Div(id="wordcloud-container", className="wordcloud-container")
])

# Callback para atualizar o gráfico com base na seleção do dropdown
@app.callback(
    Output("graph", "figure"),
    Input("dropdown-column", "value")
)
def update_graph(selected_column):
    # Cria um gráfico de barras com a coluna selecionada
    fig = px.bar(df, x=selected_column, title=f"Distribuição de {selected_column}", text_auto=True)
    return fig

# Callback para exibir a nuvem de palavras apenas se a coluna específica for selecionada
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
    else:
        return None  # Retorna None para esconder o componente quando outra coluna for selecionada

# Executa o app
if __name__ == "__main__":
    app.run(debug=True)