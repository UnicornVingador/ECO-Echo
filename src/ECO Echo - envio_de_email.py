"""
===================================================================================
📧 ENVIO DE RELATÓRIO ECO ECHO VIA OUTLOOK - VERSÃO 3.2
===================================================================================
Autor: Jonathan Barbosa
Versão: 3.2
Data: 08/02/2026

DESCRIÇÃO:
    Script para enviar automaticamente o relatório de análise de requerimentos
    via Outlook para a gerente, incluindo todos os gráficos e relatórios gerados.

    🔥 VERSÃO 3.2 - COMPATÍVEL COM SISTEMA v3.2:
    - HTML já vem com imagens embutidas em Base64 (não precisa converter)
    - Suporte à estrutura de pastas timestamped + ULTIMO
    - Novo gráfico: grafico_colaboradores_por_data.png
    - Separação por data nos relatórios
    - Detecção automática da pasta mais recente

COMPATIBILIDADE:
    ✅ Sistema de Análise v3.2 (Separação por Data)
    ✅ Sistema de Análise v3.1 (HTML com imagens embutidas)
    ✅ Estrutura: resultado_analise/YYYY-MM-DD_HH-MM-SS/ e resultado_analise/ULTIMO/

===================================================================================
"""

import os
import win32com.client
from datetime import datetime
from pathlib import Path
import glob

# ===================================================================================
# CONFIGURAÇÕES - PERSONALIZE AQUI
# ===================================================================================

class ConfigEmail:
    """Configurações do e-mail"""

    # ========== DESTINATÁRIOS ==========
    EMAIL_GERENTE = "conceicao.thiago@yduqs.com.br"
    
    # OPCIONAL: Adicionar cópia para outras pessoas
    # EMAIL_CC = "outro.email@yduqs.com.br"  # Descomente para usar
    
    # ========== CAMINHOS ==========
    # Pasta base dos relatórios (compatível com v3.2)
    PASTA_RELATORIOS_BASE = r"C:\Users\paulo.vivano\Corporativo\Experiência do Aluno Digital - CSC - Relacionamento Digital\Planilhas Tratadas\ECO Echo\resultado_analise"
    
    # Opções de pasta:
    # "ULTIMO" = sempre pega a pasta ULTIMO (mais recente)
    # "AUTO" = detecta automaticamente a pasta mais recente (timestamped)
    # "2026-02-08_15-30-45" = pasta específica
    USAR_PASTA = "ULTIMO"  # Recomendado: "ULTIMO"
    
    # ========== ARQUIVOS A ANEXAR ==========
    # Arquivos gerados pelo Sistema v3.2
    ARQUIVOS_ANEXAR = [
        "relatorio_executivo.html",           # ⭐ JÁ VEM COM IMAGENS EMBUTIDAS
        "relatorio_consolidado.xlsx"         # Excel com dados
    ]
    
    # ========== CONTEÚDO DO E-MAIL ==========
    ASSUNTO = "📊 Relatório ECO Echo - Análise de Requerimentos (v3.2) - {data}"
    
    CORPO_EMAIL = """
    <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; line-height: 1.6;">
        <p>Olá!</p>
        
        <p>Segue o relatório consolidado de análise de requerimentos do sistema <strong>ECO Echo</strong>.</p>
        
        <h3 style="color: #00B3B0; border-bottom: 2px solid #00B3B0; padding-bottom: 5px;">
            📋 Arquivos Incluídos
        </h3>
        <ul style="line-height: 1.8;">
            <li><strong>relatorio_executivo.html</strong> - Relatório interativo completo 
            </li>
            <li><strong>relatorio_consolidado.xlsx</strong> - Dados consolidados em Excel com abas separadas</li>
        </ul>
        
        <h3 style="color: #00B3B0; border-bottom: 2px solid #00B3B0; padding-bottom: 5px;">
            💡 Como Visualizar
        </h3>
        <ol style="line-height: 1.8;">
            <li><strong>Relatório HTML:</strong> Clique duas vezes em <code>relatorio_executivo.html</code> 
                para abrir no navegador - todos os gráficos já estão embutidos!</li>
            <li><strong>Análise Detalhada:</strong> Abra o Excel para filtrar e analisar dados específicos</li>
        </ol>
        
        <div style="margin-top: 30px; padding: 15px; background: #f0f8ff; border-left: 4px solid #00B3B0; border-radius: 5px;">
            <p style="margin: 0; color: #555;">
                <strong>📅 Data de geração:</strong> {data_geracao}<br>
                <strong>⏰ Hora:</strong> {hora_geracao}<br>
                <strong>🔧 Versão:</strong> Sistema de Análise v3.2
            </p>
        </div>
        
        <p style="margin-top: 20px; color: #666; font-size: 0.9em;">
            Este e-mail foi gerado automaticamente pelo Sistema ECO Echo.<br>
            Em caso de dúvidas, entre em contato com a equipe de BackOffice Acadêmico Digital.
        </p>
    </div>
    """


# ===================================================================================
# FUNÇÕES DE DETECÇÃO E VERIFICAÇÃO
# ===================================================================================

def detectar_pasta_relatorios():
    """
    Detecta automaticamente a pasta de relatórios baseado na configuração
    
    Returns:
        Path: Caminho da pasta detectada
    """
    base = Path(ConfigEmail.PASTA_RELATORIOS_BASE)
    
    if not base.exists():
        raise FileNotFoundError(f"❌ Pasta base não encontrada: {base}")
    
    # Modo ULTIMO (recomendado)
    if ConfigEmail.USAR_PASTA == "ULTIMO":
        pasta_ultimo = base / "ULTIMO"
        if pasta_ultimo.exists():
            print(f"✅ Usando pasta ULTIMO: {pasta_ultimo}")
            return pasta_ultimo
        else:
            print("⚠ Pasta ULTIMO não encontrada, tentando detectar automaticamente...")
            ConfigEmail.USAR_PASTA = "AUTO"  # Fallback
    
    # Modo AUTO (detecta pasta mais recente)
    if ConfigEmail.USAR_PASTA == "AUTO":
        pastas_timestamped = [
            d for d in base.iterdir() 
            if d.is_dir() and d.name != "ULTIMO"
        ]
        
        if not pastas_timestamped:
            raise FileNotFoundError("❌ Nenhuma pasta de relatório encontrada")
        
        # Ordenar por data de modificação (mais recente primeiro)
        pasta_mais_recente = max(pastas_timestamped, key=lambda p: p.stat().st_mtime)
        print(f"✅ Pasta mais recente detectada: {pasta_mais_recente.name}")
        return pasta_mais_recente
    
    # Modo específico (nome de pasta fornecido)
    pasta_especifica = base / ConfigEmail.USAR_PASTA
    if pasta_especifica.exists():
        print(f"✅ Usando pasta especificada: {pasta_especifica.name}")
        return pasta_especifica
    else:
        raise FileNotFoundError(f"❌ Pasta especificada não encontrada: {ConfigEmail.USAR_PASTA}")


def verificar_outlook():
    """Verifica se o Outlook está disponível"""
    try:
        win32com.client.DispatchEx("Outlook.Application")
        print("✅ Outlook detectado e pronto para uso!")
        return True
    except Exception as e:
        print("❌ ERRO: Outlook não encontrado ou não está instalado.")
        print(f"   Detalhes: {str(e)}")
        print("\n   💡 Soluções:")
        print("      1. Certifique-se de que o Microsoft Outlook está instalado")
        print("      2. Abra o Outlook manualmente primeiro")
        print("      3. Instale: pip install pywin32")
        return False


def verificar_arquivos(pasta_relatorios):
    """
    Verifica se todos os arquivos necessários existem
    
    Args:
        pasta_relatorios (Path): Pasta onde estão os relatórios
    
    Returns:
        tuple: (bool sucesso, list arquivos_encontrados, list arquivos_faltando)
    """
    print(f"\n🔍 Verificando arquivos em: {pasta_relatorios.name}")
    print("=" * 70)
    
    arquivos_encontrados = []
    arquivos_faltando = []
    
    for arquivo in ConfigEmail.ARQUIVOS_ANEXAR:
        caminho_completo = pasta_relatorios / arquivo
        
        if caminho_completo.exists():
            tamanho = caminho_completo.stat().st_size
            tamanho_mb = tamanho / (1024 * 1024)
            
            print(f"✅ {arquivo:<45} ({tamanho_mb:.2f} MB)")
            arquivos_encontrados.append(arquivo)
        else:
            print(f"❌ {arquivo:<45} NÃO ENCONTRADO")
            arquivos_faltando.append(arquivo)
    
    print("=" * 70)
    
    if arquivos_faltando:
        print(f"\n⚠ AVISO: {len(arquivos_faltando)} arquivo(s) não encontrado(s):")
        for arq in arquivos_faltando:
            print(f"   • {arq}")
        print("\n   💡 O e-mail será criado apenas com os arquivos disponíveis.")
        
        resposta = input("\n   Deseja continuar? (s/n): ").lower()
        if resposta != 's':
            return False, [], []
    
    return True, arquivos_encontrados, arquivos_faltando


def verificar_html_com_imagens(pasta_relatorios):
    """
    Verifica se o HTML tem imagens embutidas (Base64)
    
    Args:
        pasta_relatorios (Path): Pasta onde está o HTML
    
    Returns:
        bool: True se o HTML tem imagens embutidas
    """
    html_path = pasta_relatorios / "relatorio_executivo.html"
    
    if not html_path.exists():
        return False
    
    try:
        conteudo = html_path.read_text(encoding='utf-8')
        
        # Verifica se tem strings Base64 de imagem
        tem_base64 = 'data:image/png;base64,' in conteudo
        
        if tem_base64:
            print("✅ HTML com imagens embutidas detectado (Base64)")
            return True
        else:
            print("⚠ HTML não tem imagens embutidas (versão antiga?)")
            return False
            
    except Exception as e:
        print(f"⚠ Erro ao verificar HTML: {str(e)}")
        return False


# ===================================================================================
# FUNÇÃO DE CRIAÇÃO DO E-MAIL
# ===================================================================================

def criar_email(pasta_relatorios, arquivos_anexar):
    """
    Cria o e-mail no Outlook
    
    Args:
        pasta_relatorios (Path): Pasta com os arquivos
        arquivos_anexar (list): Lista de arquivos a anexar
    """
    print("\n📧 Criando e-mail no Outlook...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Destinatários
        email.To = ConfigEmail.EMAIL_GERENTE
        
        # CC (se configurado)
        if hasattr(ConfigEmail, 'EMAIL_CC'):
            email.CC = ConfigEmail.EMAIL_CC
        
        # Assunto
        email.Subject = ConfigEmail.ASSUNTO.format(
            data=datetime.now().strftime("%d/%m/%Y")
        )
        
        # Corpo do e-mail
        email.HTMLBody = ConfigEmail.CORPO_EMAIL.format(
            data_geracao=datetime.now().strftime("%d/%m/%Y"),
            hora_geracao=datetime.now().strftime("%H:%M:%S")
        )
        
        # Anexar arquivos
        print("\n📎 Anexando arquivos:")
        for arquivo in arquivos_anexar:
            caminho_arquivo = pasta_relatorios / arquivo
            email.Attachments.Add(str(caminho_arquivo.absolute()))
            print(f"   ✅ {arquivo}")
        
        # Exibir e-mail (não envia automaticamente)
        email.Display()
        
        print("\n✅ E-mail criado com sucesso!")
        print("   💡 O e-mail foi aberto no Outlook - revise e clique em ENVIAR")
        
        return True
        
    except Exception as e:
        print(f"\n❌ ERRO ao criar e-mail: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


# ===================================================================================
# FUNÇÃO PRINCIPAL
# ===================================================================================

def enviar_relatorio():
    """Função principal que orquestra todo o processo"""
    
    print("=" * 80)
    print("📧 ECO ECHO - ENVIO DE RELATÓRIO v3.2")
    print("=" * 80)
    print(f"Horário: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 80)
    
    # Etapa 1: Verificar Outlook
    print("\n[1/5] Verificando Outlook...")
    if not verificar_outlook():
        return False
    
    # Etapa 2: Detectar pasta de relatórios
    print("\n[2/5] Detectando pasta de relatórios...")
    try:
        pasta_relatorios = detectar_pasta_relatorios()
    except FileNotFoundError as e:
        print(f"\n❌ {str(e)}")
        return False
    
    # Etapa 3: Verificar se HTML tem imagens embutidas
    print("\n[3/5] Verificando formato do HTML...")
    verificar_html_com_imagens(pasta_relatorios)
    
    # Etapa 4: Verificar arquivos
    print("\n[4/5] Verificando arquivos necessários...")
    sucesso, arquivos_ok, arquivos_faltando = verificar_arquivos(pasta_relatorios)
    
    if not sucesso:
        print("\n❌ Operação cancelada pelo usuário")
        return False
    
    if not arquivos_ok:
        print("\n❌ Nenhum arquivo disponível para anexar")
        return False
    
    # Etapa 5: Criar e-mail
    print("\n[5/5] Criando e-mail no Outlook...")
    sucesso_email = criar_email(pasta_relatorios, arquivos_ok)
    
    if sucesso_email:
        print("\n" + "=" * 80)
        print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
        print("=" * 80)
        print("\n📋 Resumo:")
        print(f"   • Arquivos anexados: {len(arquivos_ok)}")
        print(f"   • Destinatário: {ConfigEmail.EMAIL_GERENTE}")
        print(f"   • Pasta: {pasta_relatorios.name}")
        
        if arquivos_faltando:
            print(f"\n   ⚠ Arquivos não encontrados (não anexados): {len(arquivos_faltando)}")
        
        print("\n💡 Próximos passos:")
        print("   1. Revise o conteúdo do e-mail no Outlook")
        print("   2. Verifique os anexos")
        print("   3. Clique em ENVIAR")
        print("=" * 80)
        return True
    else:
        print("\n❌ Falha ao criar e-mail")
        return False


# ===================================================================================
# EXECUÇÃO PRINCIPAL
# ===================================================================================

if __name__ == "__main__":
    try:
        enviar_relatorio()
    except KeyboardInterrupt:
        print("\n\n⚠ Operação cancelada pelo usuário (Ctrl+C)")
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        input("\n⏎ Pressione ENTER para fechar...")