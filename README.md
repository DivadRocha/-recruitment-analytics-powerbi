# 📊 Recruitment Analytics Dashboard - Power BI

> **Análise completa do processo de recrutamento usando Power Query + Power BI**  
> Transformando dados brutos de RH em insights acionáveis para otimização de contratação

![Power BI](https://img.shields.io/badge/Power_BI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power_Query-217346?style=for-the-badge)

---

## 📌 Contexto do Projeto

Este projeto analisa **352 posições de trabalho** e **706 candidatos** de um processo de recrutamento multi-empresa entre 2023-2026. O objetivo é identificar gargalos no funil de contratação, otimizar fontes de recrutamento e reduzir time-to-hire.

### 🎯 Objetivos de Negócio

1. **Reduzir Time-to-Hire** - Identificar etapas que atrasam o processo
2. **Otimizar ROI de Recrutamento** - Descobrir canais mais eficazes (LinkedIn vs Agências)
3. **Melhorar Taxa de Conversão** - Analisar onde perdemos candidatos qualificados
4. **Prever Necessidades Futuras** - Padrões de turnover e substituições

---

## 📊 Dataset Overview

### Tabelas Principais

| Tabela | Registros | Campos Críticos | Descrição |
|--------|-----------|-----------------|-----------|
| **POSIÇÕES** | 352 | Job ID, Company, Status, Opening/Closing Date | Vagas abertas e fechadas |
| **CANDIDATOS** | 706 | Candidate Name, Source, Recruitment Phase, Hiring Date | Pipeline de candidatos |
| **DROP DOWN** | - | Listas de validação | Empresas, Departamentos, Status |

### Qualidade dos Dados (Pré-Limpeza)

| Problema | Impacto | Registros Afetados |
|----------|---------|-------------------|
| ⚠️ Vagas sem data de fechamento | **80.7%** | 284 de 352 |
| ⚠️ Candidatos sem motivo de recusa | **97.3%** | 687 de 706 |
| ⚠️ Vagas sem candidatos | **25.6%** | 90 de 352 |
| ⚠️ Variações em SOURCE | Análise enviesada | "Linkedin" vs "LINKEDIN" |
| ⚠️ Candidatos duplicados | Contagem incorreta | 81 registros (37 únicos) |

---

## 🔧 Processo de Transformação (Power Query)

### Etapa 1: Limpeza de POSIÇÕES

```m
// Padronização de COMPANY
= Table.ReplaceValue(Source, "TM ", "TM", Replacer.ReplaceText, {"COMPANY"})
= Table.ReplaceValue(Source, "TM-IT ", "TM", Replacer.ReplaceText, {"COMPANY"})

// Categorização de TIPO DE VAGA (167 valores → 4 categorias)
= Table.AddColumn(#"Previous Step", "Tipo_Vaga_Categoria", each 
    if Text.Contains([TIPO DE VAGA], "Nova") then "Nova Vaga"
    else if Text.Contains([TIPO DE VAGA], "Substituição") then "Substituição"
    else if Text.Contains([TIPO DE VAGA], "Replecement") then "Substituição"
    else "Outros"
)

// Conversão de OPENING DATE para Date
= Table.TransformColumnTypes(#"Previous Step", {{"OPENING DATE", type date}})

// Criação de coluna calculada: Dias_Vaga_Aberta
= Table.AddColumn(#"Previous Step", "Dias_Aberta", each 
    if [CLOSING DATE] = null then 
        Duration.Days(DateTime.LocalNow() - [OPENING DATE])
    else 
        Duration.Days([CLOSING DATE] - [OPENING DATE])
)

// Tratamento de BUDGETED inconsistente
= Table.ReplaceValue(#"Previous Step", "Yes", "YES", Replacer.ReplaceText, {"BUDGETED"})
```

### Etapa 2: Limpeza de CANDIDATOS

```m
// Padronização de SOURCE (11 variações → 7 categorias)
= Table.ReplaceValue(Source, "LINKEDIN", "LinkedIn", Replacer.ReplaceText, {"SOURCE"})
= Table.ReplaceValue(#"Previous Step", "Spontaneous application", "Spontaneous Application", Replacer.ReplaceText, {"SOURCE"})
= Table.ReplaceValue(#"Previous Step", "internal recruitment", "Internal Recruitment", Replacer.ReplaceText, {"SOURCE"})

// Remoção de espaços em STATUS
= Table.TransformColumns(#"Previous Step", {{"STATUS ", Text.Trim}})

// Criação de flag: Candidato_Duplicado
= Table.AddColumn(#"Previous Step", "Flag_Duplicado", each 
    Table.RowCount(Table.SelectRows(#"Previous Step", (r) => r[CANDIDATE NAME] = [CANDIDATE NAME])) > 1
)

// Conversão de datas para type date
= Table.TransformColumnTypes(#"Previous Step", {
    {"DATE 1", type date},
    {"DATE 2", type date},
    {"DATE 3", type date},
    {"HIRING DATE", type date},
    {"Proposal Date", type date}
})
```

### Etapa 3: Criação de Tabela Dimensão - Calendário

```m
// Tabela DIM_CALENDARIO
let
    StartDate = #date(2023, 1, 1),
    EndDate = #date(2026, 12, 31),
    NumberOfDays = Duration.Days(EndDate - StartDate) + 1,
    Dates = List.Dates(StartDate, NumberOfDays, #duration(1,0,0,0)),
    DateTable = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),
    ChangedType = Table.TransformColumnTypes(DateTable, {{"Date", type date}}),
    
    // Adicionar colunas calculadas
    Ano = Table.AddColumn(ChangedType, "Ano", each Date.Year([Date])),
    Mes = Table.AddColumn(Ano, "Mês", each Date.Month([Date])),
    MesNome = Table.AddColumn(Mes, "MêsNome", each Date.MonthName([Date])),
    Trimestre = Table.AddColumn(MesNome, "Trimestre", each "Q" & Number.ToText(Date.QuarterOfYear([Date]))),
    Semana = Table.AddColumn(Trimestre, "Semana", each Date.WeekOfYear([Date]))
in
    Semana
```

---

## 📈 Métricas Chave (KPIs)

### Eficiência de Recrutamento

| Métrica | Fórmula DAX | Benchmark |
|---------|-------------|-----------|
| **Time-to-Hire Médio** | `AVERAGE(Dias_Aberta)` | < 30 dias |
| **Taxa de Preenchimento** | `DIVIDE(Filled, Total Vagas)` | > 85% |
| **Custo por Contratação** | `Total Custos / Total Hired` | Variável |

### Funil de Conversão

```
Total Candidatos (706)
    ↓ 82% (Phase 1)
Triagem/CV Aprovado (579)
    ↓ 20.4% (Phase 2)
Entrevista (118)
    ↓ 22.9% (Phase 3)
Proposta (27)
    ↓ 66.3% (Hired)
Contratados (179)
```

**Taxa de Conversão Global: 25.4%** (179 contratados / 706 candidatos)

### Performance por Canal

| SOURCE | Candidatos | Hired | Taxa Conversão | Recomendação |
|--------|------------|-------|----------------|--------------|
| LinkedIn | 425 | TBD | TBD% | 🔍 Analisar |
| Recruitment Agency | 81 | TBD | TBD% | 🔍 Analisar |
| Reference | 71 | TBD | TBD% | 🔍 Analisar |

---

## 🎨 Dashboard - Estrutura de Páginas

### 📄 Página 1: Overview Executivo

**Visuals:**
- 🔢 4 Cards: Total Vagas | Vagas Preenchidas | Time-to-Hire | Taxa Conversão
- 📊 Gráfico de Linha: Evolução Mensal de Vagas (Abertas vs Fechadas)
- 🧮 Funil: Conversão de Candidatos por Fase
- 🏢 Tabela: Top 5 Departamentos com Mais Contratações

**Filtros:**
- Company
- Ano (2023-2026)
- Tipo de Vaga (Nova vs Substituição)

---

### 📄 Página 2: Análise de Candidatos

**Visuals:**
- 🎯 Mapa de Calor: SOURCE x RECRUITMENT PHASE (conversão)
- 📊 Gráfico de Barras: Motivos de Recusa (quando disponível)
- 📈 Scatter Plot: Salary Expectation vs WAGE RANGE
- 🔢 Card: Candidatos Duplicados (flag de alerta)

**Insights Esperados:**
- Qual canal traz candidatos que chegam mais longe no funil?
- Onde estamos perdendo mais candidatos (recusa vs rejeição)?

---

### 📄 Página 3: Performance de Recrutadores

**Visuals:**
- 📊 Tabela: Recrutador | Vagas Geridas | % Fechadas | Dias Médios
- 📈 Gráfico de Dispersão: Vagas Geridas vs Taxa de Sucesso
- 🎯 Top 3 Recrutadores (Cards)

**Métricas DAX:**
```DAX
Vagas_por_Recrutador = 
    CALCULATE(
        COUNTROWS(POSIÇÕES),
        ALLEXCEPT(POSIÇÕES, POSIÇÕES[RECRUITER])
    )

Taxa_Fechamento_Recrutador = 
    DIVIDE(
        CALCULATE(COUNTROWS(POSIÇÕES), POSIÇÕES[STATUS] = "Filled"),
        COUNTROWS(POSIÇÕES)
    )
```

---

## 🚨 Problemas Identificados & Soluções Propostas

### 1. Dados de Recusa Críticos Ausentes (97.3%)

**Problema:** Apenas 19 de 706 candidatos têm motivo de recusa documentado.

**Impacto:** 
- Impossível identificar padrões de perda de talento
- Não sabemos se perdemos por salário, cultura, ou processo

**Solução Proposta:**
- Implementar campo obrigatório em sistema de ATS
- Criar categorias padronizadas de recusa
- Meta: 80% de preenchimento em 3 meses

**Métrica de Sucesso:** `Taxa_Preenchimento_Recusa > 80%`

---

### 2. Vagas Órfãs (25.6% sem candidatos)

**Problema:** 90 vagas não têm nenhum candidato associado.

**Possíveis Causas:**
- Vagas canceladas antes de recrutamento
- Problemas de integração de dados
- Vagas confidenciais/não publicadas

**Ação:** Filtrar estas vagas nas análises de conversão ou marcar como "Not Started"

---

### 3. Time-to-Hire Indefinido (80.7%)

**Problema:** 284 vagas não têm CLOSING DATE.

**Solução Power Query:**
```m
// Estimar CLOSING DATE para vagas "Filled" sem data
= Table.AddColumn(#"Previous Step", "Closing_Date_Estimado", each 
    if [STATUS] = "Filled" and [CLOSING DATE] = null then
        // Usar HIRING DATE mais recente do candidato dessa vaga
        List.Max(Table.SelectRows(CANDIDATOS, (c) => c[Job Id] = [JOB ID])[HIRING DATE])
    else
        [CLOSING DATE]
)
```

---

## 🔍 Insights Preliminares (Baseados em EDA)

### ✅ Positivos

1. **Taxa de conversão global de 25.4%** é razoável para mercado
2. **LinkedIn é canal dominante** (60% dos candidatos) - consistente com B2B pharma
3. **266 vagas preenchidas de 352** (75.6%) mostra eficácia do processo

### ⚠️ Áreas de Melhoria

1. **Funil quebra muito entre Phase 1 e 2** (82% → 20.4%) - possível problema em triagem
2. **80.7% vagas sem data de fechamento** - disciplina de dados precisa melhorar
3. **Candidatos duplicados (11.5%)** - risco de contar mesma pessoa várias vezes

---

## 🛠️ Tecnologias Utilizadas

- **Power BI Desktop** - Visualização e dashboarding
- **Power Query (M Language)** - ETL e transformação de dados
- **DAX** - Métricas e cálculos avançados
- **Excel** - Fonte de dados original

---

## 📂 Estrutura do Repositório

```
recruitment-analytics/
│
├── data/
│   ├── raw/
│   │   └── RECRUTAMENTO.xlsx          # Dados originais
│   └── processed/
│       └── recruitment_cleaned.pbix   # Modelo Power BI
│
├── docs/
│   ├── data_dictionary.md             # Dicionário de dados
│   ├── power_query_steps.md           # Transformações detalhadas
│   └── dax_measures.md                # Fórmulas DAX documentadas
│
├── images/
│   ├── dashboard_overview.png
│   ├── candidate_analysis.png
│   └── recruiter_performance.png
│
└── README.md                          # Este arquivo
```

---

## 📝 Próximos Passos

- [ ] Implementar dashboard completo em Power BI
- [ ] Criar relatório automatizado mensal (PDF export)
- [ ] Integrar com sistema ATS para dados em tempo real
- [ ] Análise preditiva: ML para prever sucesso de candidato
- [ ] Benchmark com indústria (dados públicos de RH)

---

## 👤 Autor

David Rocha 
Data Analyst | Power BI Specialist  

---

## 📄 Licença

Este projeto é para fins de portfolio e educacionais. Os dados foram anonimizados.

---

**⭐ Se este projeto foi útil, considere dar uma estrela no repositório!**
