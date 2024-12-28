import pandas as pd
import streamlit as st
from fpdf import FPDF
from streamlit_sortables import sort_items  # Adiciona suporte para ordenação manual
from PIL import Image as PILImage
import io
import folium
from geopy.geocoders import Nominatim

st.set_page_config(layout="wide",
    page_title="PEDIDOS NICOPEL",
    page_icon="🌎",
)
col1, col2 = st.columns(2)

with st.sidebar:
    st.image("img/Logo_Nicopel.png", use_column_width=True)

def estilo():
    page_bg_img = """
    <style>
    /* Alterando o fundo da área principal para preto */
    [data-testid="stAppViewContainer"] > .main {
        background-color: #09090a !important;  /* Cor de fundo preto */
    }

    /* Customizando a cor do texto do botão */
    .stButton > button {
        color: white !important;  /* Definindo a cor do texto para branco */
        background-color: #003147;  /* Cor do fundo */
        border: none;  /* Removendo bordas */
        padding: 10px 20px;  /* Ajustando o padding */
        font-size: 16px;  /* Tamanho da fonte */
    }

    /* Customizando a cor do texto do selectbox */
    div.stSelectbox > label {
        color: black !important;  /* Cor do texto */
    }

    [data-testid="stSidebar"] > div:first-child {
        background-color: #faf7f2;  /* Cor da barra lateral */
        color: black;  /* Cor do texto */
    }

    [data-testid="stHeader"] {
        background: rgba(0,0,0,0);
    }

    [data-testid="stToolbar"] {
        right: 2rem;
    }
    </style>
    """
    st.markdown(page_bg_img, unsafe_allow_html=True)

# Aplicar o estilo
estilo()

# Título estilizado
st.markdown(
    """
    <style>
    .styled-container {
        background-color:rgb(1, 68, 16); /* Verde elegante */
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        font-family: Arial, sans-serif;
        font-size: 24px;
        font-weight: bold;
        box-shadow: 0px 4px 6px rgba(12, 12, 12, 0.1);
        margin-bottom: 20px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown('<div class="styled-container">Organizador de Pedidos do Sistema ERP</div>', unsafe_allow_html=True)
def get_coordinates(row):
    geolocator = Nominatim(user_agent="myGeocoder")
    # Concatenando as informações para formar um endereço completo
    address = f"{row['Logradouro']}, {row['Bairro']}, {row['Cidade']}, {row['Número']}, {row['UF']}"
    location = geolocator.geocode(address, timeout=10)
    if location:
        return location.latitude, location.longitude
    return None, None
# Função para processar os dados do Excel
def process_excel_data(file):
    # Ler o arquivo Excel
    df = pd.read_excel(file, header=None)
    processed_data = []

    # Variáveis para armazenar informações temporárias
    current_client = ""
    current_city = ""
    current_order = ""
    current_seller = ""
    current_date = ""
    current_invoice_type = ""
    current_freight_type = ""
    current_freight_value = ""
    current_total_value = ""
    current_obs = ""

    for i, row in df.iterrows():
        # Identificar as linhas com informações do pedido
        if "//" in str(row[2]) and " - " in str(row[2]):
            # Atualiza as informações gerais do pedido
            current_order = row[0]
            current_client = row[2]
            current_city = row[3]
            current_seller = row[4]
            current_date = row[5]
            current_invoice_type = row[6]
            current_freight_type = row[7]
            current_freight_value = row[8]
            current_total_value = row[10]
            current_obs = row[11] if len(row) > 11 else ""
        
        # Identificar os itens, mas excluir a linha onde a coluna "Qtd" contém "Qtd"
        if pd.to_numeric(row[0], errors='coerce') is not None and row[3] != "" and row[1] != "Qtd":
            numero_os = row[0]
            qtd = row[1]
            descricao_item = row[3]

            processed_data.append([current_order, current_client, current_city, current_date, 
                                   current_freight_type,  qtd, descricao_item, current_obs])

    # Criar DataFrame com os dados processados
    df_processed = pd.DataFrame(processed_data, columns=["Nº Pedido",  "Cliente", "Cidade Faturamento", 
                                                         "Data Pedido", "Tipo Frete", 
                                                         "Qtd", "Descrição Item Faturamento", "Obs."])

    # Remover a primeira linha (linha 1) e as linhas com células vazias nas colunas "Qtd" e "Descrição Item Faturamento"
    df_processed = df_processed.drop(index=0)  # Remove a linha 1 (índice 0)
    df_processed = df_processed.dropna(subset=["Qtd", "Descrição Item Faturamento"])

    # Converter a coluna "Data Pedido" para datetime e criar uma coluna "Semana"
    df_processed["Data Pedido"] = pd.to_datetime(df_processed["Data Pedido"], errors='coerce')
    df_processed['Semana'] = df_processed['Data Pedido'].dt.isocalendar().week

    return df_processed
#======================================
def add_watermark(image_path, output_path, opacity=0.1):
    # Abrir a imagem original com Pillow
    image = PILImage.open(image_path).convert("RGBA")
    
    # Criar uma nova imagem com a mesma largura e altura da original
    width, height = image.size
    watermark = PILImage.new("RGBA", (width, height), (0, 0, 0, 0))

    # Colocar a imagem original na nova imagem, aplicando a opacidade
    image = image.convert("RGBA")
    for i in range(width):
        for j in range(height):
            r, g, b, a = image.getpixel((i, j))
            watermark.putpixel((i, j), (r, g, b, int(a * opacity)))

    # Salvar a nova imagem com a opacidade ajustada
    watermark.save(output_path)
#======================================
# Função para gerar o PDF
def gerar_pdf(df_filtered, frete_tipo, semana,  cidades ):
    # Caminho das imagens
    logo_path = "img/Logo_Nicopel.png"
    watermark_path = "watermarked_logo.png"
    
    # Criar o objeto FPDF com orientação 'L' para Landscape (horizontal)
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Adicionar a imagem da empresa no canto superior esquerdo
    pdf.image(logo_path, x=10, y=10, w=40)

    # Adicionar o título do relatório
    pdf.set_font("Arial", style='B', size=16)
    pdf.cell(0, 10, txt="Relatório de Pedidos", ln=True, align="C")
    pdf.ln(5)

    # Adicionar informações de filtro no cabeçalho
    pdf.set_font("Arial", size=8)
    pdf.cell(80, 5, txt=f"Semana: {semana if semana else 'Não especificado'}", ln=False)
    pdf.cell(80, 5, txt=f"Tipo de Frete: {frete_tipo if frete_tipo else 'Não especificado'}", ln=True)
    pdf.cell(80, 5, txt=f"Cidades: {cidades if cidades else 'Não especificado'}", ln=True)
    pdf.ln(5)  # Adiciona uma linha em branco entre o cabeçalho e a tabela

    # Adicionar cabeçalho da tabela (ajustado para largura da página)
    pdf.set_font("Arial", size=8)
    pdf.set_fill_color(0, 255, 0)  # Verde (RGB)

# Cabeçalho da tabela
    pdf.cell(15, 5, txt="Pedido", border=1, align="C", fill=True)  # Aplicando a cor de fundo
    pdf.cell(100, 5, txt="Cliente", border=1, align="C", fill=True)
    pdf.cell(100, 5, txt="Item", border=1, align="C", fill=True)
    pdf.cell(17, 5, txt="Qtd", border=1, align="C", fill=True)
    pdf.cell(21, 5, txt="Conferência", border=1, align="C", fill=True)  # Nova coluna "Conferência"
    pdf.cell(21, 5, txt="Faltante", border=1, align="C", fill=True)      # Nova coluna "Faltante"
    pdf.ln(5)  # Adiciona uma linha abaixo

    # Inicializar variável para o pedido anterior
    last_order = None

    # Preencher os dados da tabela com as cores e espaço entre pedidos
    for index, row in df_filtered.iterrows():
        # Adicionar um espaço extra se o número do pedido mudar
        if row['Nº Pedido'] != last_order:
            # Adicionar uma linha em branco para separar os pedidos
            pdf.ln(2)  # 5mm de espaço entre os pedidos

        # Aplique a cor no PDF com base no valor da coluna "color"
        if row['color'] == 0:
            pdf.set_fill_color(173, 216, 230)  # Azul claro
        else:
            pdf.set_fill_color(255, 255, 255)  # Branco

        # Adicionar dados do pedido com fundo colorido
        pdf.cell(15, 5, txt=str(row['Nº Pedido']), border=1, align="C", fill=True if row['color'] == 0 else False)
        pdf.cell(100, 5, txt=str(row['Cliente Nome']), border=1, align="L", fill=True if row['color'] == 0 else False)
        pdf.cell(100, 5, txt=str(row['Descrição Item Faturamento']), border=1, align="C", fill=True if row['color'] == 0 else False)
        pdf.cell(17, 5, txt=str(row['Qtd']), border=1, align="C", fill=True if row['color'] == 0 else False)
        pdf.cell(21, 5, txt="", border=1, align="C", fill=True if row['color'] == 0 else False)  # Conferência (vazio)
        pdf.cell(21, 5, txt="", border=1, align="C", fill=True if row['color'] == 0 else False)  # Faltante (vazio)
        
        # Pular uma linha após cada pedido
        pdf.ln()

        # Atualizar o número do último pedido
        last_order = row['Nº Pedido']

    # Salvar o arquivo PDF no diretório local
    pdf_output = "pedido_relatorio.pdf"  # Caminho para salvar no diretório local
    
    # Salvar o PDF
    pdf.output(pdf_output)
    return pdf_output
#======================================
# Função para o botão de download
def download_pdf(pdf_output):
    with open(pdf_output, "rb") as f:
        st.download_button("Baixar Relatório PDF", f, file_name="pedido_relatorio.pdf", mime="application/pdf")
#======================================
# Função para aplicar a cor na linha inteira
def apply_color(row):
    color = 'background-color: lightblue' if row['color'] == 0 else 'background-color: white'
    return [color] * len(row)
#======================================
with col1:
    uploaded_file = st.file_uploader("Carregue o arquivo Excel dos pedidos", type=["xlsx"])
with col2:
    uploaded_file2 = st.file_uploader("Carregue o arquivo Excel dos Cliente", type=["xlsx"])

if uploaded_file is not None:
    # Processar os dados do Excel
    df_processed = process_excel_data(uploaded_file)

    # Filtros
    st.sidebar.write("### Filtros")
    frete_types = df_processed["Tipo Frete"].unique()
    selected_frete = st.sidebar.selectbox("📌 Selecione o Tipo de Frete:", ["Todos"] + list(frete_types))

    semanas = df_processed['Semana'].unique()
    selected_semana = st.sidebar.selectbox("📌 Selecione a Semana do Pedido:", ["Todas"] + list(semanas))

    cidades = df_processed['Cidade Faturamento'].unique()
    selected_cidades = st.sidebar.multiselect("📌 Selecione as Cidades-Estado:", cidades, default=None)

    clientes = df_processed['Cliente'].unique()  # Obter clientes únicos
    excluded_clientes = st.sidebar.multiselect("📌 Excluir Clientes:", clientes, default=None)  # Multiselect para excluir clientes

    # Filtrar os dados com base nos filtros selecionados
    df_filtered = df_processed
    if selected_frete != "Todos":
        df_filtered = df_filtered[df_filtered["Tipo Frete"] == selected_frete]
    if selected_semana != "Todas":
        df_filtered = df_filtered[df_filtered["Semana"] == selected_semana]
    if selected_cidades:
        df_filtered = df_filtered[df_filtered["Cidade Faturamento"].isin(selected_cidades)]
    if excluded_clientes:
        df_filtered = df_filtered[~df_filtered["Cliente"].isin(excluded_clientes)]  # Excluir os clientes selecionados

    with st.sidebar:
        # Extrair apenas o nome do cliente após "//"
        df_filtered['Cliente Nome'] = df_filtered['Cliente'].str.split('//').str[1].str.strip()

        # Obter a lista de clientes únicos para reorganização
        clientes_filtered = df_filtered['Cliente Nome'].unique().tolist()
        st.sidebar.write("### Reorganizar Clientes")
        sorted_clientes = sort_items(clientes_filtered)

        st.markdown("***Grupo Nicopel Embalagens***")
        st.markdown('''Aplicativo desenvolvido por: ''')
        st.markdown(''':green-background[João Gabriel Brighenti]''')
        st.markdown(''':green[Todos os direitos reservados ©. V1.3.2]''')

    if sorted_clientes:
        # Aplicar filtro com base na reorganização dos clientes
        df_filtered = df_filtered[df_filtered['Cliente Nome'].isin(sorted_clientes)]

        # Ordenar pela lista reorganizada
        df_filtered['Cliente Nome'] = pd.Categorical(df_filtered['Cliente Nome'], categories=sorted_clientes, ordered=True)
        df_filtered = df_filtered.sort_values('Cliente Nome')

    # Recalcular a coluna color após os filtros e a reorganização
    df_filtered['color'] = df_filtered['Nº Pedido'].ne(df_filtered['Nº Pedido'].shift()).cumsum() % 2

    # Exibir DataFrame filtrado com altura aumentada
    st.dataframe(df_filtered.style.apply(apply_color, axis=1), hide_index=True, use_container_width=True, height=600)
    # Gerar e permitir o download do PDF

    col1, col2, col3, col4 = st.columns([1,1,1,3])

    with col1:

        motorista = st.selectbox("Selecione o Motorista", ['Ivan', 'Ronaldo'], use_container_width=True)

    with col3:
        pdf_output = gerar_pdf(df_filtered, selected_frete, selected_semana, selected_cidades)
        download_pdf(pdf_output)



if uploaded_file2 is not None:
    if uploaded_file2.name.endswith("csv"):
        # Se for CSV, carregar com pandas e especificar o cabeçalho correto
        df2 = pd.read_csv(uploaded_file2, header=0)  # Ajuste o número da linha conforme necessário
    elif uploaded_file2.name.endswith("xlsx"):
        # Se for Excel, carregar com pandas e especificar o cabeçalho correto
        df2 = pd.read_excel(uploaded_file2, header=0)  # Ajuste o número da linha conforme necessário

if uploaded_file2 is not None and not df_filtered.empty:
    # Extraindo o ID do cliente da coluna "Cliente" na primeira tabela (df_filtered)
    df_filtered['ID Cliente'] = df_filtered['Cliente'].str.split(" -").str[0]
    
    # Garantir que as colunas de IDs sejam do mesmo tipo (string ou int)
    df_filtered['ID Cliente'] = df_filtered['ID Cliente'].astype(str)  # Convertendo para string
    df2['ID'] = df2['ID'].astype(str)  # Garantindo que a coluna ID de df2 seja string também
    
    # Verificando se as colunas necessárias estão presentes em ambas as tabelas
    if "ID Cliente" in df_filtered.columns and "ID" in df2.columns:
        # Realizando a junção (merge) entre as duas tabelas com base na coluna de IDs
        df_linked = pd.merge(df_filtered, df2, left_on="ID Cliente", right_on="ID", how="left")  # Usando left join para manter todos os registros de df_filtered
        
        # Selecionando apenas as colunas desejadas
        df_selected = df_linked[['Nº Pedido', 'Cliente', 'Logradouro','Número', 'CEP', 'Cidade', 'Bairro', 'UF']]
        
        # Removendo duplicatas com base no número do pedido, mantendo apenas a primeira ocorrência
        df_selected = df_selected.drop_duplicates(subset='Nº Pedido')
        
        # Exibindo o DataFrame resultante com as colunas selecionadas
        st.write("Aqui estão os dados linkados com as colunas selecionadas (sem repetições de nº pedido):")
        df_selected = df_selected[df_selected['Logradouro'].notna()]
        #st.dataframe(df_selected, hide_index=True, use_container_width=True, height=600)
    else:
        st.error("As colunas 'ID Cliente' e 'ID' não estão presentes nas tabelas.")

#======================================
def mapa():
    
    df_selected[['Latitude', 'Longitude']] = df_selected.apply(get_coordinates, axis=1, result_type="expand")

    # Filtrando as colunas relevantes para o mapa
    df_map = df_selected[['Nº Pedido', 'Cliente', 'Logradouro', 'Número', 'CEP', 'Cidade', 'Bairro', 'UF', 'Latitude', 'Longitude']]
    #st.dataframe(df_map)
    # Remover linhas com coordenadas inválidas (NaN)
    df_map = df_map.dropna(subset=['Latitude', 'Longitude'])

    # Garantir que o DataFrame não está vazio antes de tentar criar o mapa
    if not df_map.empty:
        # Criando o mapa com um ponto central (média das coordenadas, por exemplo)
        m = folium.Map(location=[df_map['Latitude'].mean(), df_map['Longitude'].mean()], zoom_start=12)

        # Adicionando marcadores
        for idx, row in df_map.iterrows():
            folium.Marker(
                location=[row['Latitude'], row['Longitude']],
                popup=f"Pedido: {row['Nº Pedido']}<br>Cliente: {row['Cliente']}<br>Endereço: {row['Logradouro']} {row['Número']}<br>Bairro: {row['Bairro']}<br>Cidade: {row['Cidade']}<br>CEP: {row['CEP']}<br>UF: {row['UF']}",
            ).add_to(m)

        # Adicionando uma linha conectando os pontos (caminho)
        valid_coords = df_map[['Latitude', 'Longitude']].values
        if len(valid_coords) > 1:  # Garantindo que existam pelo menos 2 coordenadas válidas
            folium.PolyLine(locations=valid_coords, color='blue', weight=2.5, opacity=1).add_to(m)

        # Exibindo o mapa no Streamlit
        st.write("Mapa Interativo com Endereços")
        st.components.v1.html(m._repr_html_(), height=600)
    else:
        st.error("Não há coordenadas válidas para exibir no mapa.")


if st.button(f"Ver Mapa Iterativo", use_container_width=True, type="primary"):
    mapa()