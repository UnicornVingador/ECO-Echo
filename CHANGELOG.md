# 📝 Changelog

Todas as mudanças notáveis neste projeto serão documentadas neste arquivo.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/),
e este projeto adere ao [Semantic Versioning](https://semver.org/lang/pt-BR/).

---

## [3.2.0] - 2026-02-08

### 🔑 Adicionado - BREAKING CHANGE
- **Separação por Data**: Colaboradores agora são identificados por `NOME (DD.MM.YYYY)`
- Nova coluna `COLABORADOR_COM_DATA` no DataFrame processado
- Novo gráfico: `grafico_colaboradores_por_data.png`
- Aba "Por Colaborador (Data)" no Excel
- Badge "🔑 Separado por Data" no HTML

### ✨ Melhorado
- KPI principal agora mostra top colaborador+data (não mais agregado)
- HTML com nota explicativa sobre não-agregação de datas
- Detecção de datas únicas identificadas durante carregamento
- Console output mais detalhado sobre separação por data

### 🐛 Corrigido
- **CRÍTICO**: Versões anteriores somavam produtividade de datas diferentes incorretamente
- Análise temporal agora reflete corretamente a produtividade diária

### 📚 Documentação
- README.md completo com 15.000+ palavras
- Exemplos visuais de antes/depois da mudança
- Diagramas de fluxo de dados
- Seção de troubleshooting expandida

### 🔧 Técnico
```python
# Antes (v3.1)
df['COLABORADOR_FINAL']  # Apenas nome

# Agora (v3.2)
df['COLABORADOR_COM_DATA']  # Nome + Data
# Exemplo: "Andrey (05.02.2026)"
```

---

## [3.1.0] - 2026-02-07

### 🖼️ Adicionado
- **Imagens embutidas em Base64** no HTML
- Função `_fig_to_base64()` para converter gráficos matplotlib
- HTML autocontido (~220 KB) perfeito para email
- Dicionário `self.imagens_base64` para armazenar conversões

### ✨ Melhorado
- Relatório HTML não depende mais de arquivos PNG externos
- Email pode ser encaminhado sem perder formatação
- Processo de envio simplificado (remove etapa de conversão)

### ♻️ Refatorado
- Todos os métodos `_grafico_*()` agora salvam PNG **e** convertem para Base64
- HTML template atualizado para usar `<img src="data:image/png;base64,..."`

### 🐛 Corrigido
- HTML quebrado quando enviado por email (imagens não apareciam)
- Necessidade de anexar 5 arquivos (agora só precisa do HTML)

### 🔧 Técnico
```python
def _fig_to_base64(self, fig) -> str:
    buffer = BytesIO()
    fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    buffer.seek(0)
    image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
    buffer.close()
    plt.close(fig)
    return f"data:image/png;base64,{image_base64}"
```

---

## [3.0.0] - 2026-02-06

### 📅 Adicionado - BREAKING CHANGE
- **Data por Pasta**: Sistema usa nome da pasta (DD.MM.YYYY) como fonte temporal
- Estrutura obrigatória: `planilhas_para_analise/DD.MM.YYYY/*.xlsm`
- Nova coluna `DATA_PASTA` e `DATA_PASTA_STR`
- Busca recursiva em subpastas com `glob(recursive=True)`

### ✨ Melhorado
- Ignora coluna `DT_INICIO_ETAPA` (dados inconsistentes)
- Análise temporal mais confiável
- Detecção automática de formato DD.MM.YYYY via regex

### 🗑️ Removido
- Dependência de colunas de data dentro das planilhas
- Lógica de fallback para `DT_INICIO_ETAPA`

### 🐛 Corrigido
- **CRÍTICO**: Datas dentro das planilhas estavam incorretas/faltando
- Análise temporal era impossível com dados inconsistentes

### 🔧 Técnico
```python
def _extrair_data_da_pasta(self, caminho_arquivo: str) -> Optional[datetime]:
    pasta_pai = Path(caminho_arquivo).parent.name
    padrao = r'(\d{2})\.(\d{2})\.(\d{4})'
    match = re.search(padrao, pasta_pai)
    if match:
        dia, mes, ano = match.groups()
        return datetime(int(ano), int(mes), int(dia))
    return None
```

---

## [2.5.0] - 2026-02-01

### ✨ Adicionado
- Leitura de formatação condicional via OpenPyXL
- Classe `LeitorFormatacaoCondicional`
- Coluna `TEM_FORMATACAO` no DataFrame
- Suporte a múltiplas cores de formatação

### 🐛 Corrigido
- Pandas não detectava células formatadas condicionalmente
- Fallback robusto quando formatação não é lida

---

## [2.0.0] - 2026-01-20

### ✨ Adicionado
- Sistema de pastas timestamped (`YYYY-MM-DD_HH-MM-SS`)
- Pasta `ULTIMO/` sempre com análise mais recente
- Script separado para envio de email (`envio_de_email.py`)
- Detecção automática da pasta mais recente

### ♻️ Refatorado
- Separação de responsabilidades (análise vs envio)
- Configuração centralizada em classe `Config`

---

## [1.5.0] - 2026-01-10

### ✨ Adicionado
- Gráficos profissionais com Matplotlib/Seaborn
- Paleta de cores YDUQS
- 4 tipos de visualização (barras, evolução, pizza, top 10)
- Exportação PNG em alta resolução (300 DPI)

---

## [1.0.0] - 2025-12-15

### 🎉 Lançamento Inicial
- Leitura de planilhas Excel com Pandas
- Cálculo de KPIs básicos
- Exportação para Excel consolidado
- Análise por colaborador
- Identificação de requerimentos respondidos

### Funcionalidades Core:
- ✅ Lê múltiplas planilhas
- ✅ Detecta colaborador pelo nome do arquivo
- ✅ Conta requerimentos respondidos
- ✅ Gera Excel consolidado

---

## 🚀 Roadmap Futuro

### [3.3.0] - Planejado Q1 2026
- [ ] Suporte a Google Sheets via API
- [ ] Dashboard interativo com Plotly
- [ ] Exportação automática para PowerPoint
- [ ] Envio programado (scheduler)
- [ ] Notificações no Microsoft Teams

### [4.0.0] - Visão Q2 2026
- [ ] Interface gráfica (GUI) com PyQt6
- [ ] Banco de dados PostgreSQL para histórico
- [ ] Machine Learning para previsão de prazos
- [ ] API REST para integração com outros sistemas
- [ ] Autenticação e controle de acesso

---

## 📋 Convenções

### Tipos de Mudança
- `✨ Adicionado` - Novas funcionalidades
- `✨ Melhorado` - Melhorias em funcionalidades existentes
- `🐛 Corrigido` - Correções de bugs
- `♻️ Refatorado` - Mudanças de código sem alterar comportamento
- `🗑️ Removido` - Funcionalidades removidas
- `🔒 Segurança` - Correções de vulnerabilidades
- `📚 Documentação` - Apenas mudanças em documentação
- `🔧 Técnico` - Detalhes de implementação

### Breaking Changes
Marcados com **BREAKING CHANGE** e explicação do impacto.

---

## 🔗 Links Úteis

- [Guia de Contribuição](CONTRIBUTING.md) (em desenvolvimento)
- [Documentação Completa](README.md)
- [Issues](https://github.com/seu-usuario/eco-echo/issues)
- [Releases](https://github.com/seu-usuario/eco-echo/releases)

---

**Última atualização:** 08/02/2026  
**Mantido por:** Jonathan Barbosa - YDUQS CSC
