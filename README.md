Descrição do Código
Este aplicativo possui as seguintes funcionalidades:

Upload de Arquivos:

Permite o upload de dois arquivos:
Um arquivo Excel contendo informações dos pedidos.
Um arquivo Excel ou CSV contendo informações dos clientes.
Processamento de Dados:

Extrai e processa informações relevantes dos pedidos, como número do pedido, cliente, cidade, data, tipo de frete, quantidade e descrição dos itens.
Remove dados duplicados e vazios, organiza por semana e inclui colunas derivadas, como "Semana" (baseada na data do pedido).
Interface de Filtros:

Oferece filtros para refinar os dados exibidos:
Tipo de frete.
Semana do pedido.
Cidade de faturamento.
Inclui uma funcionalidade para reorganizar a lista de clientes manualmente.
Exibição de Dados:

Exibe uma tabela interativa com os dados processados e filtrados.
Aplica formatação condicional às linhas para melhorar a visualização.
Geração de Relatório em PDF:

Cria um relatório em PDF com os pedidos filtrados.
Inclui informações como número do pedido, cliente, descrição do item, quantidade, e campos extras ("Conferência" e "Faltante").
Mapeamento de Localizações:

Utiliza a biblioteca geopy para obter as coordenadas geográficas (latitude e longitude) com base nos endereços dos clientes.
Gera um mapa interativo com os locais dos pedidos utilizando a biblioteca folium.
Estilo e Personalização:

Personaliza o layout do Streamlit, como cores de fundo, botões estilizados, e barras laterais.
Inclui uma marca d'água em imagens usadas no aplicativo.
Download de Relatórios:

Permite aos usuários baixar o relatório PDF gerado.
Integração de Dados:

Faz a junção (merge) entre os pedidos e as informações dos clientes para exibir dados combinados, como endereço e cidade.
Organização Visual:

Inclui um cabeçalho estilizado e apresenta uma mensagem de autoria e versão na barra lateral.
Objetivo do Aplicativo
Este aplicativo foi desenvolvido para facilitar o gerenciamento de pedidos no ambiente de uma empresa que utiliza sistemas ERP. Ele permite:

Centralizar informações de pedidos e clientes.
Automatizar a análise de dados, reduzindo o tempo necessário para processar arquivos manualmente.
Gerar relatórios personalizados de forma rápida e eficiente.
Fornecer insights geográficos sobre os pedidos com mapas interativos.
Melhorar a visualização e organização dos pedidos para tomada de decisões mais informada.
Tecnologias Utilizadas
Streamlit: Framework principal para a interface do usuário.
Pandas: Manipulação de dados.
FPDF: Geração de relatórios em PDF.
folium: Criação de mapas interativos.
geopy: Obtenção de coordenadas geográficas.
Pillow: Processamento de imagens.
Possíveis Casos de Uso
Empresas que precisam organizar e acompanhar pedidos de forma dinâmica.
Equipes de vendas ou logística que precisam de relatórios personalizados.
Análise e visualização geográfica de clientes e pedidos.
Como Executar
Clone este repositório.
Instale as dependências necessárias listadas no arquivo requirements.txt.
Execute o aplicativo com o comando:
bash
Copiar código
streamlit run app.py
Faça o upload dos arquivos Excel ou CSV de pedidos e clientes para começar a usar.