# Painel da Agropecuária Municipal do Brasil, 2025

Painel interativo com a área, a quantidade e o valor da produção agrícola, o efetivo dos rebanhos e o crédito rural contratado em cada município brasileiro em 2025, a partir da Produção Agrícola Municipal (PAM) e da Pesquisa da Pecuária Municipal (PPM) do IBGE e da Matriz de Dados do Crédito Rural do Banco Central. Elaboração: Instituto Arapyaú.

Ao selecionar um bioma ou um estado no tema Produção, aparece uma tabela com os produtos mais produzidos (agricultura, por valor da produção) ou com os rebanhos (pecuária, por número de cabeças) do recorte.

O painel tem três temas, **Produção**, **Crédito rural** e **Imóveis rurais (CAR)**, e dois setores, **Agricultura** e **Pecuária** (no crédito, também **Todos**). Os filtros de bioma e estado valem para todos.

## Estrutura

```
index.html                 página do painel (mapa, filtros, totais e ranking)
incorporar.html            código para incorporar o painel em outro site
dados/pam_2025.json        agricultura, no formato lido pela página
dados/pam_2025.csv         agricultura em CSV, para download
dados/ppm_2025.json        pecuária, no formato lido pela página
dados/ppm_2025.csv         pecuária em CSV, para download
dados/credito_2025.json    crédito rural, no formato lido pela página
dados/credito_2025.csv     crédito rural em CSV, para download
dados/credito_produtos_2025.json  crédito por produto financiado e estado, no formato lido pela página
dados/credito_produtos_2025.csv   a mesma base em CSV, para download
dados/instituicoes_financeiras.csv  nome e tipo de cada instituição financeira (CNPJ de 8 dígitos)
dados/malha.json           contornos dos municípios e estados, já projetados
dados/pecuaria_produtos_uf_2025.csv  origem animal e aquicultura por estado (IBGE, PPM)
dados/car_2025.json         imóveis rurais no CAR (cadastros e área), no formato lido pela página
dados/car_2025.csv          a mesma base em CSV, para download
dados/biomas_municipios.csv  bioma predominante de cada município (IBGE)
dados/malha_biomas.json    contornos dos biomas, já projetados
vendor/d3.v7.9.0.min.js    biblioteca de gráficos (D3)
vendor/fonts/              fonte Open Sans (Regular e ExtraBold)
scripts/consolidar.py      gera os arquivos de dados/ a partir das planilhas do IBGE
scripts/municipios_ibge.csv  lista oficial de municípios com código IBGE
```

A página não depende de nenhum serviço externo: biblioteca, fontes e dados estão todos no repositório.

## Como publicar no GitHub Pages

1. Crie um repositório **público** (por exemplo, `pam-2025-arapyau`).
2. Em **Add file → Upload files**, arraste todo o conteúdo desta pasta (mantendo as subpastas) e clique em **Commit changes**.
3. Em **Settings → Pages**, escolha **Deploy from a branch**, branch `main`, pasta `/ (root)`, e salve.
4. Em um ou dois minutos, o painel estará em `https://<sua-conta>.github.io/pam-2025-arapyau/`.

Para atualizar uma versão já publicada, envie os arquivos novos pelo mesmo caminho: arquivos com o mesmo nome substituem os antigos.

Para testar no computador antes de publicar, abra um terminal nesta pasta, rode `python -m http.server` e acesse `http://localhost:8000`. Abrir o `index.html` com duplo clique não funciona, porque o navegador bloqueia a leitura dos dados fora de um servidor.

## Links para uma visão específica

Os filtros ficam gravados no endereço da página. Exemplos:

- `…/#tipo=Permanente&produto=Açaí&variavel=valor&uf=PA`: valor da produção de açaí no Pará.
- `…/#setor=pecuaria&rebanho=Todos&vista=predominante`: rebanho predominante em cada município.
- `…/#tema=credito&setor=agricultura&programa=Pronaf&linha=Floresta (sistemas agroflorestais)&bioma=Amazônia`: Pronaf Floresta na Amazônia.
- `…/#tema=credito&setor=todos&instituicao=Cooperativa de crédito`: crédito rural contratado em cooperativas de crédito.
- `…/#setor=pecuaria&pecVer=aquicultura&pecProd=Tilápia`: produção de tilápia por estado.
- `…/#tema=car&carTam=Grande&carVar=pct_area&bioma=Amazônia`: participação das grandes propriedades na área cadastrada no CAR, na Amazônia.
- `…/#tema=credito&setor=todos&produtoCr=Cacau`: crédito rural para cacau, por estado.
- `…/#tema=credito&setor=todos&uf=PA`: crédito rural no Pará, com o quadro dos produtos mais financiados no estado.
- `…/#tema=credito&setor=todos&ifin=04902979`: crédito rural contratado com o Banco da Amazônia (o código é o CNPJ de 8 dígitos da instituição, que aparece no endereço ao clicar nela no quadro "Maiores instituições financeiras").
- `…/#bioma=Amazônia&tipo=Permanente&produto=Cacau (em amêndoa)&variavel=valor`: valor do cacau nos municípios da Amazônia.

## Como atualizar os dados

Quando o IBGE publicar uma revisão:

1. Baixe as planilhas do SIDRA, nível município, um arquivo por UF: Tabela 1612 (lavouras temporárias, arquivos `Tabela_3_*`), Tabela 1613 (lavouras permanentes, `Tabela_6_*`) e Tabela 3939 (rebanhos, `Tabela_2_*`). Para o crédito rural, baixe a Matriz de Dados do Crédito Rural do Banco Central (arquivo `MATRIZ_CR_REGIAO_*.csv`) e a matriz por produto, base por estado (`MATRIZ_CR_PRODUTO_REGIAO_*.csv`). Coloque tudo em uma pasta.
2. Rode `pip install openpyxl` e depois `python scripts/consolidar.py CAMINHO/DA/PASTA`. O script processa os arquivos que encontrar na pasta: dá para atualizar só a agricultura, só a pecuária ou só o crédito.
3. Suba os novos arquivos de `dados/` no repositório, substituindo os antigos.

Se o script avisar de municípios sem correspondência, inclua a grafia em `CORRECOES` no início do script.

Para Minas Gerais, as lavouras temporárias desta versão vieram da Tabela 5457 baixada por município (arquivos `Minas_Gerais_-_area_e_quantidade.xlsx` e `Minas_Gerais_-_Valor.xlsx`), porque a Tabela 3.11 não estava disponível. O script aceita esse formato: qualquer planilha da Tabela 5457 por município colocada na mesma pasta é usada para as UFs que não tiverem arquivo `Tabela_3_*`.

## Fonte e método

- **Agricultura:** IBGE, PAM 2025, Tabelas 1612 e 1613 (para as temporárias de Minas Gerais, Tabela 5457 por município, versão consolidada da mesma pesquisa). Área plantada (temporárias) ou área destinada à colheita (permanentes); em "Todas as lavouras", as duas são somadas. Quantidade só por produto, em toneladas (abacaxi e coco-da-baía em mil frutos). Valor em mil reais, preço pago ao produtor.
- **Pecuária:** IBGE, PPM 2025, Tabela 3939, efetivo em 31/12, em cabeças. Matrizes de suínos e galinhas são subgrupos de suínos e galináceos.
- **Origem animal e aquicultura:** IBGE, PPM 2025, Tabelas 74 e 3940, por estado (no painel, em Produção → Pecuária → Ver). Nessas visões o mapa mostra estados e o filtro de bioma não aparece. As quantidades são exibidas em toneladas, litros, dúzias ou unidades. Para atualizar, coloque as planilhas das duas tabelas (por estado) na pasta do script.
- **Todos os rebanhos:** soma das cabeças de todos os rebanhos do município, sem os subgrupos (matrizes de suínos e galinhas), que já estão nos totais. O rebanho predominante é o grupo com mais cabeças no município; como as aves são muito mais numerosas, predominam em muitos municípios.
- **Crédito rural:** Banco Central, Matriz de Dados do Crédito Rural (SICOR), operações contratadas de janeiro a dezembro de 2025 pela data de emissão (não o ano-safra), no município informado na operação, em valores nominais. Agregado por município, programa, linha (Pronaf e RenovAgro), setor, finalidade e instituição financeira (nome, CNPJ de 8 dígitos e tipo); fonte de recursos e mês ficaram de fora. Programas não listados em `PROGRAMAS` no script são avisados na execução.
- **Produtos financiados:** Matriz de Dados do Crédito Rural por produto, base por estado (completa). A base por município não é usada porque o Banco Central omite o município das combinações com até 3 contratos (sigilo bancário), o que deixa sem município 49% do valor, de forma desigual entre produtos e programas. Por isso, o recorte por produto mostra estados e não aceita o filtro de bioma.
- **Imóveis rurais (CAR):** Ministério do Meio Ambiente e Mudança do Clima, Painel de Regularização Ambiental, extração de outubro de 2026: número de cadastros e área cadastrada (declarada) por município, tamanho (pequena até 4, média de 4 a 15 e grande acima de 15 módulos fiscais), tipo de imóvel (IRU, AST, PCT) e condição do CAR. O tema abre com os imóveis rurais (IRU), porque assentamentos e territórios tradicionais costumam ser cadastrados como um único imóvel coletivo na faixa de até 4 módulos fiscais. Pode haver sobreposição entre cadastros, e a área cadastrada não equivale à área ocupada. Para atualizar, coloque as duas planilhas exportadas do Painel (cadastros e área) na pasta do script.
- **Biomas:** bioma predominante por município, IBGE (lista de 2024, com base na Malha Municipal 2022). Municípios na transição entre biomas entram inteiros no bioma que ocupa a maior parte do território; o filtro é um recorte por município, não uma medição exata dentro de cada bioma. Para atualizar, substitua `dados/biomas_municipios.csv` mantendo as colunas `cod_ibge`, `municipio`, `uf` e `bioma`.
- **Resultados preliminares:** o IBGE classifica 2025 como preliminar, sujeito a revisão.
- **Malha municipal:** projeto [geodata-br](https://github.com/tbrugz/geodata-br), simplificada e projetada (Albers). Base anterior a 2013: seis municípios criados depois não têm polígono próprio, mas seus dados entram nos totais e no ranking.
- **Códigos IBGE:** atribuídos pelo nome do município e UF, com a lista do projeto [municipios-brasileiros](https://github.com/kelvins/municipios-brasileiros). Dezesseis grafias foram ajustadas (ver `scripts/consolidar.py`).
