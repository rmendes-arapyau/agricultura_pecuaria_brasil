# Produção agrícola e pecuária municipal do Brasil, 2025

Painel interativo com a área, a quantidade e o valor da produção agrícola e o efetivo dos rebanhos de cada município brasileiro, a partir da Produção Agrícola Municipal (PAM) e da Pesquisa da Pecuária Municipal (PPM) 2025 do IBGE. Elaboração: Instituto Arapyaú.

## Estrutura

```
index.html                 página do painel (mapa, filtros, totais e ranking)
incorporar.html            código para incorporar o painel em outro site
dados/pam_2025.json        agricultura, no formato lido pela página
dados/pam_2025.csv         agricultura em CSV, para download
dados/ppm_2025.json        pecuária, no formato lido pela página
dados/ppm_2025.csv         pecuária em CSV, para download
dados/malha.json           contornos dos municípios e estados, já projetados
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
- `…/#modo=pecuaria&rebanho=Todos&vista=predominante`: rebanho predominante em cada município.
- `…/#bioma=Amazônia&tipo=Permanente&produto=Cacau (em amêndoa)&variavel=valor`: valor do cacau nos municípios da Amazônia.

## Como atualizar os dados

Quando o IBGE publicar uma revisão:

1. Baixe as planilhas do SIDRA, nível município, um arquivo por UF: Tabela 1612 (lavouras temporárias, arquivos `Tabela_3_*`), Tabela 1613 (lavouras permanentes, `Tabela_6_*`) e Tabela 3939 (rebanhos, `Tabela_2_*`). Coloque todas em uma pasta.
2. Rode `pip install openpyxl` e depois `python scripts/consolidar.py CAMINHO/DA/PASTA`. O script processa as tabelas que encontrar na pasta: dá para atualizar só a agricultura ou só a pecuária.
3. Suba os novos arquivos de `dados/` no repositório, substituindo os antigos.

Se o script avisar de municípios sem correspondência, inclua a grafia em `CORRECOES` no início do script.

Para Minas Gerais, as lavouras temporárias desta versão vieram da Tabela 5457 baixada por município (arquivos `Minas_Gerais_-_area_e_quantidade.xlsx` e `Minas_Gerais_-_Valor.xlsx`), porque a Tabela 3.11 não estava disponível. O script aceita esse formato: qualquer planilha da Tabela 5457 por município colocada na mesma pasta é usada para as UFs que não tiverem arquivo `Tabela_3_*`.

## Fonte e método

- **Agricultura:** IBGE, PAM 2025, Tabelas 1612 e 1613 (para as temporárias de Minas Gerais, Tabela 5457 por município, versão consolidada da mesma pesquisa). Área plantada (temporárias) ou área destinada à colheita (permanentes); em "Todas as lavouras", as duas são somadas. Quantidade só por produto, em toneladas (abacaxi e coco-da-baía em mil frutos). Valor em mil reais, preço pago ao produtor.
- **Pecuária:** IBGE, PPM 2025, Tabela 3939, efetivo em 31/12, em cabeças. Matrizes de suínos e galinhas são subgrupos de suínos e galináceos.
- **Todos os rebanhos:** efetivo convertido em unidades animais com os coeficientes da FAO para a América do Sul (FAO, 2011, *Guidelines for the preparation of livestock sector reviews*, Anexo 1): bovino 0,70; suíno 0,25; equino 0,65; ovino e caprino 0,10; galináceos 0,01. Búfalo (0,70) e codorna (0,01) por aproximação, pois a FAO não traz coeficiente para eles na região. O rebanho predominante é o grupo com mais unidades animais no município.
- **Biomas:** bioma predominante por município, IBGE (lista de 2024, com base na Malha Municipal 2022). Municípios na transição entre biomas entram inteiros no bioma que ocupa a maior parte do território; o filtro é um recorte por município, não uma medição exata dentro de cada bioma. Para atualizar, substitua `dados/biomas_municipios.csv` mantendo as colunas `cod_ibge`, `municipio`, `uf` e `bioma`.
- **Resultados preliminares:** o IBGE classifica 2025 como preliminar, sujeito a revisão.
- **Malha municipal:** projeto [geodata-br](https://github.com/tbrugz/geodata-br), simplificada e projetada (Albers). Base anterior a 2013: seis municípios criados depois não têm polígono próprio, mas seus dados entram nos totais e no ranking.
- **Códigos IBGE:** atribuídos pelo nome do município e UF, com a lista do projeto [municipios-brasileiros](https://github.com/kelvins/municipios-brasileiros). Dezesseis grafias foram ajustadas (ver `scripts/consolidar.py`).
