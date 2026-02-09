# ⚡ Guia Rápido - ECO Echo

**Comece a usar o sistema em 5 minutos!**

---

## 📦 Instalação Rápida

### Passo 1: Clone o projeto
```bash
git clone https://github.com/seu-usuario/eco-echo.git
cd eco-echo
```

### Passo 2: Instale as dependências
```bash
pip install -r requirements.txt
```

### Passo 3: Organize suas planilhas

Crie a estrutura de pastas:
```
planilhas_para_analise/
├── 05.02.2026/
│   ├── gestao_requerimentos_andrey.xlsm
│   └── gestao_requerimentos_maria.xlsm
└── 06.02.2026/
    └── gestao_requerimentos_andrey.xlsm
```

**⚠️ IMPORTANTE:** Nome da pasta deve ser `DD.MM.YYYY`

---

## 🚀 Uso Básico

### Gerar Relatório

```bash
python "ECO Echo2.py"
```

**Resultado:** Pasta `resultado_analise/ULTIMO/` com 6 arquivos

### Enviar por Email

```bash
python "ECO Echo - envio_de_email.py"
```

**Resultado:** Email aberto no Outlook pronto para enviar

---

## 📋 Checklist Pré-Execução

- [ ] ✅ Python 3.8+ instalado
- [ ] ✅ Dependências instaladas (`pip install -r requirements.txt`)
- [ ] ✅ Planilhas organizadas em `planilhas_para_analise/DD.MM.YYYY/`
- [ ] ✅ Planilhas têm aba "BASE"
- [ ] ✅ Coluna I (ou "SITUAÇÃO") existe
- [ ] ✅ Valores como "Deferido", "Indeferido" na coluna

**Para envio de email:**
- [ ] ✅ Microsoft Outlook instalado
- [ ] ✅ Email configurado em `ConfigEmail.EMAIL_GERENTE`

---

## 🎯 Estrutura Mínima da Planilha

Sua planilha precisa ter:

| Coluna | Nome Sugerido | Obrigatório? | Conteúdo |
|--------|---------------|--------------|----------|
| A-H | Variáveis | Não | Dados gerais |
| **I** | **SITUAÇÃO** | **SIM** | Deferido/Indeferido/Redirecionado |
| J+ | Variáveis | Não | Outros dados |

**Aba obrigatória:** "BASE"

---

## 🔧 Configurações Básicas

### Em `ECO Echo2.py`:

```python
class Config:
    # Onde estão suas planilhas
    PASTA_PLANILHAS = "./planilhas_para_analise"
    
    # Onde salvar os resultados
    PASTA_SAIDA = "./resultado_analise"
    
    # Valores que indicam "respondido"
    VALORES_RESPONDIDO = [
        "Deferido",
        "Indeferido", 
        "Redirecionado"
    ]
```

### Em `ECO Echo - envio_de_email.py`:

```python
class ConfigEmail:
    # Email do destinatário
    EMAIL_GERENTE = "seu.chefe@empresa.com.br"
    
    # Onde estão os relatórios
    PASTA_RELATORIOS_BASE = r"C:\caminho\para\resultado_analise"
    
    # Qual pasta usar (ULTIMO é recomendado)
    USAR_PASTA = "ULTIMO"
```

---

## 🎨 Outputs Gerados

Após executar `ECO Echo2.py`, você terá:

```
resultado_analise/ULTIMO/
├── 📄 relatorio_executivo.html          (Abra no navegador)
├── 📊 relatorio_consolidado.xlsx        (Abra no Excel)
├── 📈 grafico_colaboradores_por_data.png
├── 📉 grafico_evolucao_temporal.png
├── 🥧 grafico_distribuicao_situacao.png
└── 📊 grafico_top_tipos.png
```

**Tamanho total:** ~2-5 MB

---

## 🐛 Problemas Comuns

### Erro: "Nenhuma planilha encontrada"

**Causa:** Caminho errado ou estrutura de pastas incorreta

**Solução:**
1. Verifique `Config.PASTA_PLANILHAS`
2. Certifique-se da estrutura `DD.MM.YYYY/arquivo.xlsm`

---

### Erro: "Coluna de situação não encontrada"

**Causa:** Planilha não tem coluna I ou "SITUAÇÃO"

**Solução:**
1. Abra a planilha no Excel
2. Verifique se a coluna I existe
3. Adicione valores como "Deferido" nas células

---

### Erro: "Outlook não encontrado"

**Causa:** Microsoft Outlook não está instalado ou não foi aberto

**Solução:**
1. Instale o Microsoft Outlook
2. Abra o Outlook manualmente pelo menos uma vez
3. Execute o script novamente

---

### Nenhum requerimento respondido detectado

**Causa:** Formatação condicional não está aplicada ou valores incorretos

**Solução:**
1. Verifique se `Config.VALORES_RESPONDIDO` bate com seus dados
2. Abra a planilha e confirme os valores da coluna I
3. Sistema tem fallback automático (deve funcionar mesmo sem formatação)

---

## 📚 Próximos Passos

Agora que você rodou com sucesso:

1. 📖 Leia o [README completo](README.md) para entender todas as funcionalidades
2. 🎨 Personalize as cores em `Config.CORES_PADRAO`
3. 📧 Configure o envio automático de email
4. 🔄 Agende execuções automáticas (Task Scheduler do Windows)
5. 🚀 Explore funcionalidades avançadas

---

## 💡 Dicas Pro

### Dica 1: Use caminhos absolutos
```python
PASTA_PLANILHAS = r"C:\Users\seu.nome\Desktop\planilhas"
```

### Dica 2: Teste com poucos arquivos primeiro
Coloque 2-3 planilhas para testar antes de processar tudo.

### Dica 3: Verifique os logs
Console mostra exatamente o que está sendo processado.

### Dica 4: Backup automático
Pastas timestamped servem como backup histórico.

### Dica 5: Valide o HTML
Abra `relatorio_executivo.html` antes de enviar por email.

---

## 🆘 Precisa de Ajuda?

- 📖 [README Completo](README.md)
- 📝 [Changelog](CHANGELOG.md)
- 🐛 [Reportar Bug](https://github.com/seu-usuario/eco-echo/issues)
- 💬 [Discussões](https://github.com/seu-usuario/eco-echo/discussions)

---

## ⏱️ Tempo de Execução Esperado

| Quantidade de Planilhas | Tempo Médio | Notas |
|-------------------------|-------------|-------|
| 1-5 planilhas | ~10 segundos | Rápido |
| 6-15 planilhas | ~30 segundos | Normal |
| 16-30 planilhas | ~1-2 minutos | Aceitável |
| 30+ planilhas | ~3-5 minutos | Otimize se possível |

---

## 🎯 Resultado Final

**Antes do ECO Echo:**
- ⏰ 2-3 horas de trabalho manual
- 😰 Propenso a erros
- 📊 Gráficos feitos no Excel

**Com o ECO Echo:**
- ⚡ 30 segundos automático
- ✅ 100% preciso
- 🎨 Relatórios profissionais

---

<div align="center">

**⭐ Pronto para começar? Execute agora:**

```bash
python "ECO Echo2.py"
```

[⬆ Voltar ao README](README.md)

</div>
