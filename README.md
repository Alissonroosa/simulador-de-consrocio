# Simulador de Consórcios – Sicredi

Aplicação estática (HTML + JS + CSS) que lê as bases em CSV da pasta `data/`. Não tem build nem servidor próprio.

**Fluxo:** tela de dados → **Simular** → tela de resultados. Na tela de resultados, **Editar simulação** volta ao formulário com os dados preenchidos e **Nova simulação** limpa o formulário. Cada clique em Simular gera um número (`SIM-AAAAMMDD-XXXXX`) e pode registrar a simulação numa lista do SharePoint (veja abaixo).

```
simulador-consorcio/
├─ index.html
├─ index.js
├─ style.css
├─ assets/logo-sicredi.png      ← opcional: logo positivo baixado em marca.sicredi.com.br
└─ data/
   ├─ planos.csv                  (obrigatório)
   ├─ indexadores.csv             (obrigatório)
   ├─ historico_lances.csv
   ├─ financiamento.csv
   ├─ produtos_investimento.csv
   ├─ produtos_credito_lance.csv
   └─ parametros.csv
```

## Publicação no SharePoint

1. Suba a pasta inteira para uma biblioteca do site (ex.: `Site Assets/simulador-consorcio/`).
2. O SharePoint Online baixa arquivos `.html` em vez de exibi-los. Renomeie `index.html` para **`index.aspx`**.
3. Na página, adicione a web part **Incorporar (Embed)** com:
   ```html
   <iframe src="https://SEU-TENANT.sharepoint.com/sites/SEU-SITE/SiteAssets/simulador-consorcio/index.aspx"
           width="100%" height="2200" style="border:0"></iframe>
   ```
4. Se a web part injetar o HTML direto na página (ex.: Modern Script Editor), informe URLs absolutas no elemento raiz:
   `data-path="https://.../simulador-consorcio/data/"` e `data-logo="https://.../assets/logo-sicredi.png"`.

Para testar localmente, rode `python -m http.server` na pasta e abra `http://localhost:8000`. Abrindo com duplo clique (`file://`) o navegador bloqueia a leitura dos CSVs.

**Scripts personalizados:** dependendo da configuração do tenant, o upload de arquivos `.aspx` pode estar bloqueado. Nesse caso, fale com o administrador do SharePoint. O caminho oficialmente suportado é empacotar estes mesmos arquivos numa web part SPFx.

**Bibliotecas externas:** a exportação de PDF usa jsPDF e jsPDF-AutoTable (cdnjs), e as fontes vêm do Google Fonts. Se a rede bloquear esses domínios, copie os `.js` para a biblioteca e ajuste o `src` no `index.html`. Sem o jsPDF, o botão abre a impressão do navegador.

## Registro das simulações

### Como funciona
O app roda dentro do SharePoint com o usuário já autenticado. Ao clicar em **Simular**, ele grava um item numa **lista do SharePoint** pela API REST do próprio site:
- não precisa de licença premium;
- a coluna padrão *Criado por* identifica o usuário.

O **Power Automate** entra depois, com o gatilho *"Quando um item é criado"* nessa lista. Exemplos de uso: notificar o gerente, copiar para Excel ou Dataverse, alimentar um painel no Power BI.

**Condição:** a página precisa estar no mesmo domínio do SharePoint (`seutenant.sharepoint.com`), como no `.aspx` da biblioteca ou numa web part que injeta o HTML. Se o iframe apontar para outro domínio, a gravação na lista não funciona.

### Configuração
1. Crie uma lista (ex.: **Simulações Consórcio**) com as colunas abaixo.
   - Crie cada coluna **exatamente com esse nome**, sem acento e sem espaço, para que o nome interno fique igual. Depois você pode renomear o nome de exibição.
   - Colunas que não existirem na lista são ignoradas.
2. Em `data/parametros.csv`, preencha `registro_lista_sharepoint` com o nome da lista. Deixe `registro_site_url` vazio se a lista estiver no mesmo site do app.
3. Permissões: os usuários precisam poder **adicionar itens** na lista. Para que cada um veja só as próprias simulações, use *Configurações da lista → Configurações avançadas → Acesso a itens*: ler e editar apenas os itens criados pelo usuário.

| Coluna (nome interno) | Tipo | Conteúdo |
|---|---|---|
| Title (Título) | Texto | nº da simulação `SIM-...` |
| IdOrigem | Texto | simulação que foi editada para gerar esta |
| Associado | Texto | |
| Consultor | Texto | |
| Segmento | Texto | Imóveis, Automóveis… |
| Plano | Texto | código do plano |
| PrazoMeses | Número | |
| ValorCredito | Moeda | |
| TaxaAdmPct | Número | |
| FundoReservaPct | Número | |
| Indice | Texto | INCC, IPCA… |
| TipoParcela | Texto | Integral / Reduzida |
| ReducaoPct | Número | |
| LanceProprioPct | Número | |
| LanceEmbutidoPct | Número | |
| ModoContemplacao | Texto | Projeção / Manual |
| MesContemplacao | Número | |
| Parcela1 | Moeda | 1ª parcela |
| ParcelaPosContemplacao | Moeda | |
| TotalDesembolsado | Moeda | |
| CustoEfetivoAA | Número | % a.a. |
| ChanceContemplacao12m | Número | % |
| LanceRecomendadoPct | Número | lance moderado na 1ª assembleia |
| DataSimulacao | Data e hora | |
| DadosJson | Várias linhas de texto (texto simples) | todas as entradas e o resumo em JSON |

Na tela de resultados aparece "Registrada · item N". Se a gravação falhar, aparece "Não registrada": passe o mouse sobre o aviso para ver o motivo, que também é gravado no console do navegador.

**Alternativa (não recomendada):** `registro_flow_url` envia o mesmo JSON, com nome e e-mail do usuário, para um fluxo com gatilho *"Quando uma solicitação HTTP é recebida"*. Esse gatilho exige licença premium, e a URL fica visível no código da página.

## Formato dos CSVs

- Separador `;` e decimal com vírgula (padrão do Excel pt-BR). Arquivos separados por `,` com decimal `.` também são aceitos.
- Salve como "CSV UTF-8" ou "CSV (separado por vírgulas)" do Excel. As duas codificações são reconhecidas.
- Não use separador de milhar (`300000,00`, não `300.000,00`).
- Os nomes das colunas não diferenciam maiúsculas nem acentos. A ordem das colunas não importa.
- Datas: `AAAA-MM-DD`, `DD/MM/AAAA` ou `AAAA-MM`.

### planos.csv (uma linha por plano)
O usuário escolhe **segmento** e **prazo**. Se houver mais de uma linha com o mesmo segmento e prazo, cada uma vale para uma faixa de crédito, e o app escolhe a linha pelo valor informado. Isso permite taxa de administração diferente por faixa.

| coluna | descrição |
|---|---|
| codigo_plano | identificador (vai para o PDF e para a lista) |
| segmento | `IMOVEL`, `AUTO`, `SERVICOS`, `PESADOS` (outros valores também funcionam) |
| descricao | texto exibido |
| prazo_meses | prazo do plano |
| taxa_adm_pct / fundo_reserva_pct | percentuais totais do plano |
| indice_reajuste | `INCC`, `IPCA`, `IGPM` |
| seguro_mensal_pct | seguro prestamista, % ao mês sobre o saldo devedor |
| taxa_adesao_pct | antecipação da taxa de administração, cobrada na 1ª parcela |
| credito_min / credito_max | faixa de crédito da linha |
| lance_embutido_max_pct | % máximo do crédito usado como lance embutido |
| lance_fixo_pct | % do lance fixo (0 se não houver) |
| permite_parcela_reduzida | `S` / `N` |
| reducao_parcela_pct | % de redução padrão (ex.: 50) |
| base_lance | `CREDITO` ou `SALDO_DEVEDOR`: sobre o que o % de lance é calculado |
| ativo | `S` / `N`: linhas com `N` não aparecem |

### historico_lances.csv (uma linha por grupo × assembleia)
`segmento; prazo_plano; grupo; assembleia; data_assembleia; cotas_aptas; contemplados_sorteio; contemplados_lance_livre; contemplados_lance_fixo; ofertas_lance_livre; ofertas_lance_fixo; menor_lance_livre_pct; maior_lance_livre_pct; media_lance_livre_pct`

O histórico alimenta a projeção de contemplação e a recomendação de lance:
- **Janela:** usa os últimos `janela_historico_meses` (padrão 24).
- **Recorte:** usa só os grupos do mesmo segmento **e prazo** quando houver pelo menos `minimo_assembleias_plano` assembleias. Caso contrário, usa o segmento inteiro.
- **Grupos:** `grupo` e `assembleia` são informativos. A simulação não pede o grupo.

### indexadores.csv (uma linha por mês, taxas **mensais** em %)
`competencia; tipo; ipca_pct; incc_pct; igpm_pct; cdi_pct; tr_pct`

`tipo` = `REALIZADO` ou `PROJETADO`. Para meses sem dado, o app repete a média dos últimos 12 meses informados.

### financiamento.csv
`codigo; segmento; modalidade; sistema_padrao (SAC|PRICE); taxa_juros_aa_pct; indexador (TR|IPCA|PRE); seguro_mip_mensal_pct; seguro_dfi_mensal_pct; tarifa_mensal; tarifa_contratacao; iof_pct; entrada_min_pct; prazo_max_meses`

### produtos_investimento.csv
`codigo; produto; indexador (CDI|IPCA|TR|PRE); percentual_indexador; taxa_adicional_aa_pct; isento_ir (S|N); liquidez; aplicacao_minima`

### produtos_credito_lance.csv
`codigo; produto; taxa_am_pct; prazo_max_meses; limite_pct_credito; iof_pct; garantia; segmentos`

`segmentos` = `TODOS` ou lista separada por `|` (ex.: `IMOVEL|PESADOS`).

### parametros.csv (`chave; valor; descricao`)
- Identificação: título e nome da cooperativa.
- Histórico: janela de meses e mínimo de assembleias do plano.
- Projeção e lance: probabilidade-alvo do mês projetado, confiança dos cenários de lance (90/70/50%) e horizonte da tendência do lance.
- Comparativos: % do CDI usado no valor presente.
- Registro: lista, site e URL do fluxo.
- Texto do aviso legal.

## Premissas de cálculo

- **Parcela** = (fundo comum + taxa de adm. + fundo de reserva) ÷ prazo, em % do crédito, × crédito atualizado, + seguro sobre o saldo devedor.
- **Reajuste:** a cada 12 meses, o crédito é corrigido pelo índice do plano acumulado nos 12 meses anteriores. Depois da contemplação, o saldo devedor continua sendo corrigido.
- **Parcela reduzida:** a redução vale até a contemplação (ou até a metade do plano, se vier antes). O valor não pago é diluído nas parcelas restantes.
- **Lance:** abate o saldo devedor. "Reduzir prazo" mantém a parcela, "reduzir parcela" mantém o prazo. O lance embutido sai do crédito recebido.
- **Contemplação:** a chance mensal combina três fontes:
  - sorteio: sorteados ÷ cotas aptas no histórico, crescendo à medida que o grupo contempla;
  - lance livre: probabilidade de o lance superar o menor lance vencedor projetado. A projeção é uma tendência linear no tempo sobre a janela e fica estável após o horizonte;
  - lance fixo: taxa histórica de contemplados ÷ ofertantes.
- **Recomendação de lance** = menor lance projetado + z × desvio histórico, onde z corresponde à confiança de cada cenário.

Os CSVs que acompanham o projeto são **dados fictícios de modelo**. Substitua pelos dados reais antes de usar com associados.
