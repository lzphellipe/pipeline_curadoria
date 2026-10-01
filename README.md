# Pipeline de curadoria de anotações LC-MS/MS — *Vanilla planifolia* (Orchidaceae)

Recebe a lista bruta de compostos anotados por bibliotecas espectrais (MassBank, NIST, MoNA, HMDB), limpa os nomes, padroniza, elimina duplicatas e classifica cada composto pela plausibilidade de ocorrer na espécie ou na família botânica.

## Como rodar

```bash
pip install -r requirements.txt
python pipeline.py dados/Lista_de_componentes_29.09.xlsx -o saida
```

Opções: `--especie "Vanilla planifolia"` e `--familia Orchidaceae` (só mudam títulos e alertas).

A entrada é uma planilha `.xlsx` com **coluna A = Número** e **coluna B = Componente**, a partir da linha 2.

## Estrutura

```
pipeline_curadoria/
├── pipeline.py                 # o pipeline (6 etapas)
├── requirements.txt
├── dados/
│   ├── Lista_de_componentes_29.09.xlsx   # entrada original
│   ├── base_curadoria.csv      # base de conhecimento: 1 linha por composto curado
│   ├── sinonimos.csv           # nome limpo  →  composto padronizado
│   └── referencias.csv         # referências bibliográficas (R1…R12)
└── saida/                      # gerado pelo pipeline
    ├── Curadoria_compostos_Vanilla_planifolia.xlsx
    ├── 2_limpo.csv             # 1 linha por linha da entrada
    ├── 3_curadoria.csv         # 1 linha por composto único
    └── pendentes.csv           # compostos sem curadoria (categoria P)
```

## Etapas

| # | Etapa | O que faz | Automática? |
|---|---|---|---|
| 1 | Carregar | Lê a planilha bruta e preserva o original | Sim |
| 2 | Limpar | Separa biblioteca, ID do banco e energia de colisão do nome (regex); detecta problemas de qualidade: encoding corrompido, espaços, caixa alta, sinônimos concatenados, aduto embutido, código no lugar do nome | Sim |
| 3 | Padronizar | Converte o nome limpo em composto canônico via `sinonimos.csv` e marca duplicatas | Sim, a partir do dicionário |
| 4 | Curar | Cruza com `base_curadoria.csv`: classe, categoria A–F, grupo isomérico, nível MSI, justificativa e referências | Sim, a partir da base |
| 5 | Verificar | Avisa quando os marcadores esperados da espécie (vanilina, glucovanilina, ácido vanílico…) não aparecem e gera os demais alertas | Sim |
| 6 | Exportar | Planilha com fórmulas vivas (o resumo recalcula se a categoria for alterada) e CSVs | Sim |

### O que foi manual (e precisa ser declarado no trabalho)

A **curadoria quimiotaxonômica**, ou seja, decidir a categoria de cada composto, foi feita por **pesquisa na literatura**, e o resultado está registrado em `dados/base_curadoria.csv`. O pipeline aplica essa base de forma reprodutível, mas não "descobre" sozinho se um composto ocorre em Orchidaceae. Fontes consultadas: metaboloma publicado de *V. planifolia* (R1, R4, R6, R7), estudos LC-MS de *Vanilla* spp. brasileiras e de outras orquídeas (R2, R3), aroma da baunilha curada (R5) e guias de contaminantes de LC-MS (R8, R9).

O mesmo vale para o `sinonimos.csv`. Nele estão decisões como "psicose/sorbose são hexoses isoméricas" e "as linhas 3 e 76 são a mesma molécula".

## Categorias

| Cat. | Significado | Ação |
|---|---|---|
| A | Relatado na espécie | Manter |
| B | Relatado no gênero ou em outras Orchidaceae | Manter |
| C | Metabólito primário/ubíquo em plantas | Manter |
| D | Possível pela biossíntese, sem registro encontrado na família | Confirmar com padrão |
| E | Improvável na família (outro táxon, microrganismo ou animal) | Excluir ou reavaliar |
| F | Contaminante, sintético, fármaco ou artefato | Excluir |
| P | Pendente: composto novo, ainda não curado | Curar manualmente |

## Usando com uma lista nova

1. Rode o pipeline com a nova lista.
2. Abra `saida/pendentes.csv`. Ele mostra os compostos que ainda não estão na base.
3. Para cada pendente:
   - se for sinônimo de um composto já curado, acrescente uma linha em `sinonimos.csv` (`nome_limpo` em minúsculas → `composto`);
   - se for um composto novo, pesquise a ocorrência (LOTUS: lotus.naturalproducts.net; KNApSAcK; literatura) e acrescente uma linha em `base_curadoria.csv`.
4. Rode de novo. O número de pendentes deve chegar a zero.

## Resultado na lista de 29/09

148 linhas → 108 compostos únicos. Havia 40 duplicatas, e 123 linhas tinham algum problema de qualidade.
Por categoria: A 14 · B 5 · C 34 · D 21 · E 14 · F 20.

## Limitações

- Todas as anotações são por similaridade espectral, portanto no máximo nível 2 do MSI. Sem m/z, tempo de retenção e score, não dá para validar as anotações nem separar isômeros.
- A busca na literatura não foi exaustiva. "Sem registro encontrado" não significa "ausente na família".
