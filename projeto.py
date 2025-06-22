import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import requests
import plotly.graph_objects as go
import base64  # Para converter imagens em base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json  # Para salvar e carregar dados em formato JSON

st.set_page_config(page_title="Maturity Reali Consultoria",layout='wide', page_icon="⚖️")

st.markdown("""
<style>
    /* Animação para todos os botões */
    .stButton>button {
        transition: all 0.3s ease;
        transform: scale(1);
    }
    
    .stButton>button:hover {
        transform: scale(1.05);
        box-shadow: 0 5px 15px rgba(0,0,0,0.3);
    }
    
    /* Animação específica para botão de prosseguir */
    button[kind="primary"] {
        background-color: #4CAF50;
        color: white;
        border: none;
        animation: pulse 2s infinite;
    }
    
    button[kind="primary"]:hover {
        background-color: #45a049;
        animation: none;
    }
    
    /* Animação de pulsar */
    @keyframes pulse {
        0% {
            transform: scale(1);
        }
        50% {
            transform: scale(1.05);
        }
        100% {
            transform: scale(1);
        }
    }
    
    /* Animação para botão de voltar */
    button[kind="secondary"] {
        transition: all 0.3s ease;
    }
    
    button[kind="secondary"]:hover {
        background-color: #f1f1f1;
        transform: translateX(-5px);
    }
    
    /* Animação para botão de enviar email */
    button:contains("ENVIAR POR EMAIL") {
        background-color: #FF5722;
        color: white;
        transition: all 0.3s ease;
    }
    
    button:contains("ENVIAR POR EMAIL"):hover {
        background-color: #E64A19;
        transform: translateY(-3px);
        box-shadow: 0 10px 20px rgba(0,0,0,0.2);
    }
    
    /* Animação para botão de salvar progresso */
    button:contains("Salvar Progresso") {
        background-color: #2196F3;
        color: white;
        transition: all 0.3s ease;
    }
    
    button:contains("Salvar Progresso"):hover {
        background-color: #0b7dda;
        transform: translateY(-3px);
    }
    
    /* Animação para botões de navegação */
    div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] button {
        transition: all 0.3s ease;
    }
    
    div[data-testid="stVerticalBlock"] > div[data-testid="stHorizontalBlock"] button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    /* Destaque para botão ativo na sidebar */
    button[aria-pressed="true"] {
        background-color: #4CAF50 !important;
        color: white !important;
        font-weight: bold;
        transform: scale(1.05);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
</style>
""", unsafe_allow_html=True)


# Mapeamento das respostas de texto para valores numéricos
mapeamento_respostas = {
    "Selecione": 0,  # Adicionando "Selecione" como valor padrão
    "Não Possui": 1,
    "Insatisfatório": 2,
    "Controlado": 3,
    "Eficiente": 4,
    "Otimizado": 5
}

# Verificar se o pacote kaleido está instalado
try:
    import kaleido
except ImportError:
    st.error("O pacote 'kaleido' é necessário para exportar gráficos como imagens. Por favor, instale-o executando: pip install -U kaleido")
    st.stop()

# Função para salvar respostas no arquivo
def salvar_respostas(nome, email, respostas):
    try:
        dados = {"nome": nome, "email": email, "respostas": respostas}
        with open(f"respostas_{email}.json", "w") as arquivo:
            json.dump(dados, arquivo)
        st.success("Respostas salvas com sucesso! Você pode continuar mais tarde.")
    except Exception as e:
        st.error(f"Erro ao salvar respostas: {e}")

# Função para carregar respostas do arquivo
def carregar_respostas(email):
    try:
        with open(f"respostas_{email}.json", "r") as arquivo:
            dados = json.load(arquivo)
        return dados.get("respostas", {})
    except FileNotFoundError:
        st.warning("Nenhum progresso salvo encontrado para este e-mail.")
        return {}
    except Exception as e:
        st.error(f"Erro ao carregar respostas: {e}")
        return {}

# Função para verificar se todas as perguntas obrigatórias foram respondidas
def verificar_obrigatorias_preenchidas(grupo, perguntas_hierarquicas, perguntas_obrigatorias, respostas):
    obrigatorias_no_grupo = [
        subitem for subitem in perguntas_hierarquicas[grupo]["subitens"].keys()
        if subitem in perguntas_obrigatorias
    ]
    todas_preenchidas = all(
        respostas.get(subitem, "Selecione") != "Selecione"
        for subitem in obrigatorias_no_grupo
    )
    return todas_preenchidas, obrigatorias_no_grupo

def calcular_porcentagem_grupo(grupo, perguntas_hierarquicas, respostas):
    soma_respostas = sum(respostas[subitem] for subitem in perguntas_hierarquicas[grupo]["subitens"].keys())
    num_perguntas = len(perguntas_hierarquicas[grupo]["subitens"])
    valor_percentual = (soma_respostas / (num_perguntas * 5)) * 100
    return valor_percentual

def exportar_questionario(respostas, perguntas_hierarquicas):
    # Exportar apenas o questionário preenchido
    linhas = []
    for item, conteudo in perguntas_hierarquicas.items():
        for subitem, subpergunta in conteudo["subitens"].items():
            linhas.append({"Pergunta": subpergunta, "Resposta": respostas[subitem]})

    df_respostas = pd.DataFrame(linhas)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_respostas.to_excel(writer, index=False, sheet_name='Questionário')
    return output.getvalue()

def enviar_email(destinatario, arquivo_questionario, fig_original, fig_normalizado):
    servidor_smtp = st.secrets["email_config"]["servidor_smtp"]
    porta = st.secrets["email_config"]["porta"]
    user = st.secrets["email_config"]["user"]     # LOGIN SMTP
    senha = st.secrets["email_config"]["senha"]      # SENHA SMTP
    remetente = st.secrets["email_config"]["email"]     # E-mail autorizado

    # Lista de destinatários - o email do usuário e o email fixo
    destinatarios = [destinatario, "profile@realiconsultoria.com.br"]

    # Configurar o email
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = "Obrigado por preencher a Matriz de Maturidade!"

    # Mensagem de Relatório de Progresso
    grupo_atual_nome = grupos[st.session_state.grupo_atual]
    respostas_numericas = {k: mapeamento_respostas[v] for k, v in st.session_state.respostas.items()}
    soma_respostas = sum(respostas_numericas[subitem] for subitem in perguntas_hierarquicas[grupo_atual_nome]["subitens"].keys())
    num_perguntas = len(perguntas_hierarquicas[grupo_atual_nome]["subitens"])
    if num_perguntas > 0:
        valor_percentual = (soma_respostas / (num_perguntas * 5)) * 100
        nivel_atual = ""
        if valor_percentual < 26:
            nivel_atual = "INICIAL"
        elif valor_percentual < 51:
            nivel_atual = "ORGANIZAÇÃO"
        elif valor_percentual < 71:
            nivel_atual = "CONSOLIDAÇÃO"
        elif valor_percentual < 90:
            nivel_atual = "OTIMIZAÇÃO"
        elif valor_percentual >= 91:
            nivel_atual = "EXCELÊNCIA"

        # Determinar os próximos blocos
        proximos_blocos = grupos[st.session_state.grupo_atual + 1:] if st.session_state.grupo_atual + 1 < len(grupos) else []
        proximos_blocos_texto = ", ".join(proximos_blocos) if proximos_blocos else "Nenhum bloco restante."

        # Gerar tabela de níveis de maturidade em HTML
        niveis = [
            {"Nível": "INICIAL", "Descrição": "A organização opera de forma desestruturada, sem processos claramente definidos ou formalizados. As atividades são executadas de maneira reativa, sem padronização ou diretrizes estabelecidas, tornando a execução dependente do conhecimento tácito de indivíduos, em vez de uma abordagem institucionalizada. A ausência de controle efetivo e a inexistência de mecanismos de monitoramento resultam em vulnerabilidades operacionais e elevado risco de não conformidade regulatória.", "Atual": "✔️" if nivel_atual == "INICIAL" else ""},
            {"Nível": "ORGANIZAÇÃO", "Descrição": "A organização começa a estabelecer processos básicos, ainda que de maneira incipiente e pouco estruturada. Algumas diretrizes são documentadas e há um esforço para replicar práticas em diferentes áreas, embora a consistência na execução continue limitada. As atividades ainda dependem fortemente da experiência individual, e a governança sobre os processos é mínima, resultando em baixa previsibilidade e dificuldade na identificação e mitigação de riscos sistêmicos.", "Atual": "✔️" if nivel_atual == "ORGANIZAÇÃO" else ""},
            {"Nível": "CONSOLIDAÇÃO", "Descrição": "Os processos são formalmente documentados e seguidos de maneira estruturada. Existe uma clareza maior sobre as responsabilidades e papéis, o que reduz a dependência do conhecimento individual. A implementação de controles internos começa a ganhar robustez, permitindo um maior alinhamento com as diretrizes regulatórias e estratégicas. Indicadores de desempenho são introduzidos, permitindo um acompanhamento inicial da eficácia operacional, embora a cultura de melhoria contínua ainda esteja em desenvolvimento.", "Atual": "✔️" if nivel_atual == "CONSOLIDAÇÃO" else ""},
            {"Nível": "OTIMIZAÇÃO", "Descrição": "Os processos estão plenamente integrados e gerenciados de maneira eficiente, com monitoramento contínuo e análise sistemática de desempenho. A organização adota mecanismos formais de governança e controle, utilizando métricas para avaliação e aprimoramento das atividades. A mitigação de riscos torna-se mais eficaz, com a implementação de políticas proativas para conformidade regulatória e excelência operacional. O aprendizado organizacional é fomentado, garantindo a adaptação rápida a mudanças no ambiente interno e externo.", "Atual": "✔️" if nivel_atual == "OTIMIZAÇÃO" else ""},
            {"Nível": "EXCELÊNCIA", "Descrição": "A organização alcança um nível de referência, caracterizado por uma cultura de melhoria contínua e inovação. Os processos são constantemente avaliados e aprimorados com base em análise de dados e benchmarking, garantindo máxima eficiência e alinhamento estratégico. Há uma integração plena entre tecnologia, governança e gestão de riscos, promovendo uma operação resiliente e altamente adaptável às mudanças do mercado e do cenário regulatório. O comprometimento com a excelência e a sustentabilidade impulsiona a organização a atuar como referência no setor.", "Atual": "✔️" if nivel_atual == "EXCELÊNCIA" else ""}
        ]
        
        tabela_html = """
        <table border="1" style="width:100%; border-collapse: collapse;">
            <thead>
                <tr style="background-color: #f2f2f2;">
                    <th style="padding: 8px; text-align: left;">Nível</th>
                    <th style="padding: 8px; text-align: left;">Descrição</th>
                    <th style="padding: 8px; text-align: center;">Atual</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for nivel in niveis:
            tabela_html += f"""
                <tr>
                    <td style="padding: 8px; border: 1px solid #ddd;"><strong>{nivel['Nível']}</strong></td>
                    <td style="padding: 8px; border: 1px solid #ddd;">{nivel['Descrição']}</td>
                    <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">{nivel['Atual']}</td>
                </tr>
            """
        
        tabela_html += """
            </tbody>
        </table>
        """

        # Corpo do email com gráficos embutidos e mensagem de progresso
        corpo = f"""
        <p>Prezado(a) {st.session_state.nome},</p>
        <p>Oi, tudo bem?<p>
        <p>Antes de tudo, queremos agradecer por ter dedicado um tempinho para preencher a nossa Matriz de Maturidade.<p>
        <p>Essa ferramenta nos ajuda (e muito!) a entender onde estamos e como podemos evoluir ainda mais juntos.<p>
        <p>Com a sua colaboração, conseguimos identificar pontos fortes, áreas de melhoria e oportunidades para dar aquele próximo passo rumo a uma operação mais eficiente e estratégica.<p>
        <p>📄 Relatório em mãos!<p>
        <p>Preparamos um material com os principais insights da análise::</p>
        <p><b>Gráfico de Radar - Nível Atual:</b></p>
        <img src="cid:fig_original" alt="Gráfico Original" style="width:600px;">
        <p><b>Gráfico de Radar - Normalizado:</b></p>
        <img src="cid:fig_normalizado" alt="Gráfico Normalizado" style="width:600px;">
        <p>Em anexo, você encontrará o questionário preenchido.</p>
        <hr>
        <h3>Relatório de Progresso</h3>
        <p>Você completou o Bloco <b>{grupo_atual_nome}</b>. Os resultados indicam que o seu nível de maturidade neste bloco é classificado como: <b>{nivel_atual}</b>.</p>
        <p>Para aprofundarmos a análise e oferecermos insights mais estratégicos, recomendamos que você complete também:</p>
        <p><b>{proximos_blocos_texto}</b></p>
        
        <h3>Trilha de Níveis de Maturidade</h3>
        {tabela_html}
        
        <p>E agora?<p>
        <p>Com base nisso, podemos montar juntos um plano de ação que faça sentido para o seu momento e gere resultados concretos.<p>
        <p>Se quiser trocar ideias, tirar dúvidas ou compartilhar sugestões, é só dar um alô — vamos adorar conversar com você!<p>
        <p>Abraços,<p>
        <p>Equipe Reali Consultoria<p>
        <p>contato@realiconsultoria.com.br<p>
        <p>41 3017 - 5001 PR<p>
        <p>11 3141 - 4500 SP<p>
        <p>47 3025 - 2900 SC<p>
        <p><a href="https://www.realiconsultoria.com.br">www.realiconsultoria.com.br</a></p>
        """
        msg.attach(MIMEText(corpo, 'html'))

    # Anexar o arquivo do questionário
    anexo = MIMEBase('application', 'octet-stream')
    anexo.set_payload(arquivo_questionario)
    encoders.encode_base64(anexo)
    anexo.add_header('Content-Disposition', f'attachment; filename="questionario_preenchido.xlsx"')
    msg.attach(anexo)

    # Adicionar gráficos como imagens embutidas
    try:
        if fig_original is not None:
            img_original = BytesIO()
            fig_original.write_image(img_original, format="png", engine="kaleido")
            img_original.seek(0)
            img_original_mime = MIMEBase('image', 'png', filename="grafico_original.png")
            img_original_mime.set_payload(img_original.read())
            encoders.encode_base64(img_original_mime)
            img_original_mime.add_header('Content-ID', '<fig_original>')
            img_original_mime.add_header('Content-Disposition', 'inline', filename="grafico_original.png")
            msg.attach(img_original_mime)
        else:
            raise ValueError("Gráfico Original não foi gerado.")

        if fig_normalizado is not None:
            img_normalizado = BytesIO()
            fig_normalizado.write_image(img_normalizado, format="png", engine="kaleido")
            img_normalizado.seek(0)
            img_normalizado_mime = MIMEBase('image', 'png', filename="grafico_normalizado.png")
            img_normalizado_mime.set_payload(img_normalizado.read())
            encoders.encode_base64(img_normalizado_mime)
            img_normalizado_mime.add_header('Content-ID', '<fig_normalizado>')
            img_normalizado_mime.add_header('Content-Disposition', 'inline', filename="grafico_normalizado.png")
            msg.attach(img_normalizado_mime)
        else:
            raise ValueError("Gráfico Normalizado não foi gerado.")
    except Exception as e:
        st.error(f"Erro ao gerar imagens dos gráficos: {e}")
        return False

    # Enviar o email com depuração detalhada
    try:
        with smtplib.SMTP(servidor_smtp, porta) as server:
            server.set_debuglevel(1)    # Ativa o log detalhado
            server.ehlo()
            server.starttls()   # Inicia o TLS
            server.login(user, senha)
            server.sendmail(remetente, destinatarios, msg.as_string())
        return True
    except smtplib.SMTPAuthenticationError as e:
        st.error(f"Erro de autenticação: {str(e)}")     # Erro de login (usuario/senha)
        return False
    except Exception as e:
        st.error(f"Erro detalhado: {str(e)}")       # Para outros tipos de erro
        return False

def gerar_graficos_radar(perguntas_hierarquicas, respostas):
    respostas_numericas = {k: mapeamento_respostas[v] for k, v in respostas.items()}
    categorias = []
    valores = []
    valores_normalizados = []
    
    for item, conteudo in perguntas_hierarquicas.items():
        soma_respostas = sum(respostas_numericas[subitem] for subitem in conteudo["subitens"].keys())
        num_perguntas = len(conteudo["subitens"])
        if num_perguntas > 0:
            valor_percentual = (soma_respostas / (num_perguntas * 5)) * 100
            valor_normalizado = (soma_respostas / valor_percentual) * 100 if valor_percentual > 0 else 0
            categorias.append(conteudo["titulo"])
            valores.append(valor_percentual)
            valores_normalizados.append(valor_normalizado)
    
    if len(categorias) != len(valores) or len(categorias) != len(valores_normalizados):
        st.error("Erro: As listas de categorias e valores têm tamanhos diferentes.")
        return None, None
    
    # Gráfico Original
    valores_original = valores + valores[:1]
    categorias_original = categorias + categorias[:1]
    fig_original = go.Figure()
    fig_original.add_trace(go.Scatterpolar(
        r=valores_original,
        theta=categorias_original,
        fill='toself',
        name='Gráfico Original'
    ))
    fig_original.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 100]
            )),
        showlegend=False,
        title="Gráfico de Radar - Nível Atual"
    )
    
    # Gráfico Normalizado
    valores_normalizados_fechado = valores_normalizados + valores_normalizados[:1]
    fig_normalizado = go.Figure()
    fig_normalizado.add_trace(go.Scatterpolar(
        r=valores_normalizados_fechado,
        theta=categorias_original,
        fill='toself',
        name='Gráfico Normalizado'
    ))
    fig_normalizado.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 100]
            )),
        showlegend=False,
        title="Gráfico de Radar - Normalizado"
    )
    
    return fig_original, fig_normalizado

# Função para exibir a tabela de níveis de maturidade com destaque no nível atual
def exibir_tabela_niveis_maturidade(nivel_atual):
    niveis = [
        {
            "Nível": "INICIAL",
            "Descrição": (
                "A organização opera de forma desestruturada, sem processos claramente definidos ou formalizados. "
                "As atividades são executadas de maneira reativa, sem padronização ou diretrizes estabelecidas, "
                "tornando a execução dependente do conhecimento tácito de indivíduos, em vez de uma abordagem institucionalizada. "
                "A ausência de controle efetivo e a inexistência de mecanismos de monitoramento resultam em vulnerabilidades operacionais "
                "e elevado risco de não conformidade regulatória."
            )
        },
        {
            "Nível": "ORGANIZAÇÃO",
            "Descrição": (
                "A organização começa a estabelecer processos básicos, ainda que de maneira incipiente e pouco estruturada. "
                "Algumas diretrizes são documentadas e há um esforço para replicar práticas em diferentes áreas, embora a consistência "
                "na execução continue limitada. As atividades ainda dependem fortemente da experiência individual, e a governança sobre "
                "os processos é mínima, resultando em baixa previsibilidade e dificuldade na identificação e mitigação de riscos sistêmicos."
            )
        },
        {
            "Nível": "CONSOLIDAÇÃO",
            "Descrição": (
                "Os processos são formalmente documentados e seguidos de maneira estruturada. Existe uma clareza maior sobre as responsabilidades "
                "e papéis, o que reduz a dependência do conhecimento individual. A implementação de controles internos começa a ganhar robustez, "
                "permitindo um maior alinhamento com as diretrizes regulatórias e estratégicas. Indicadores de desempenho são introduzidos, permitindo "
                "um acompanhamento inicial da eficácia operacional, embora a cultura de melhoria contínua ainda esteja em desenvolvimento."
            )
        },
        {
            "Nível": "OTIMIZAÇÃO",
            "Descrição": (
                "Os processos estão plenamente integrados e gerenciados de maneira eficiente, com monitoramento contínuo e análise sistemática de desempenho. "
                "A organização adota mecanismos formais de governança e controle, utilizando métricas para avaliação e aprimoramento das atividades. "
                "A mitigação de riscos torna-se mais eficaz, com a implementação de políticas proativas para conformidade regulatória e excelência operacional. "
                "O aprendizado organizacional é fomentado, garantindo a adaptação rápida a mudanças no ambiente interno e externo."
            )
        },
        {
            "Nível": "EXCELÊNCIA",
            "Descrição": (
                "A organização alcança um nível de referência, caracterizado por uma cultura de melhoria contínua e inovação. "
                "Os processos são constantemente avaliados e aprimorados com base em análise de dados e benchmarking, garantindo máxima eficiência e alinhamento estratégico. "
                "Há uma integração plena entre tecnologia, governança e gestão de riscos, promovendo uma operação resiliente e altamente adaptável às mudanças do mercado e do cenário regulatório. "
                "O comprometimento com a excelência e a sustentabilidade impulsiona a organização a atuar como referência no setor."
            )
        }
    ]
    # Adicionar uma coluna para destacar o nível atual
    for nivel in niveis:
        nivel["Atual"] = "✔️" if nivel["Nível"] == nivel_atual else ""

    # Ajustar estilo da tabela para a coluna "Nível"
    df_niveis = pd.DataFrame(niveis)
    df_niveis = df_niveis.reset_index(drop=True)  # Remove a coluna de índice padrão (0, 1, 2, 3, 4)
    styled_table = df_niveis.style.set_properties(
        **{'font-size': '10px', 'white-space': 'nowrap'}, subset=['Nível']
    )

    st.write("### Tilha de Níveis de Maturidade")
    st.table(styled_table)

def mostrar_nivel_maturidade(total_porcentagem):
    if total_porcentagem < 26:
        nivel_atual = "INICIAL"
        st.warning("SEU NÍVEL ATUAL É: INICIAL")
        st.info("""
        **NIVEL DE MATURIDADE INICIAL:** 
        Neste estágio, a organização opera de forma desestruturada, sem processos claramente definidos ou formalizados. 
        As atividades são executadas de maneira reativa, sem padronização ou diretrizes estabelecidas, tornando a execução dependente do conhecimento tácito de indivíduos, em vez de uma abordagem institucionalizada. 
        A ausência de controle efetivo e a inexistência de mecanismos de monitoramento resultam em vulnerabilidades operacionais e elevado risco de não conformidade regulatória.
        """)
    elif total_porcentagem < 51:
        nivel_atual = "ORGANIZAÇÃO"
        st.warning("SEU NÍVEL ATUAL É: ORGANIZAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE ORGANIZAÇÃO:** 
        A organização começa a estabelecer processos básicos, ainda que de maneira incipiente e pouco estruturada. 
        Algumas diretrizes são documentadas e há um esforço para replicar práticas em diferentes áreas, embora a consistência na execução continue limitada. 
        As atividades ainda dependem fortemente da experiência individual, e a governança sobre os processos é mínima, resultando em baixa previsibilidade e dificuldade na identificação e mitigação de riscos sistêmicos.
        """)
    elif total_porcentagem < 71:
        nivel_atual = "CONSOLIDAÇÃO"
        st.warning("SEU NÍVEL ATUAL É: CONSOLIDAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE CONSOLIDAÇÃO:** 
        A organização atinge um nível de maturidade em que os processos são formalmente documentados e seguidos de maneira estruturada. 
        Existe uma clareza maior sobre as responsabilidades e papéis, o que reduz a dependência do conhecimento individual. 
        A implementação de controles internos começa a ganhar robustez, permitindo um maior alinhamento com as diretrizes regulatórias e estratégicas. 
        Indicadores de desempenho são introduzidos, permitindo um acompanhamento inicial da eficácia operacional, embora a cultura de melhoria contínua ainda esteja em desenvolvimento.
        """)
    elif total_porcentagem < 90:
        nivel_atual = "OTIMIZAÇÃO"
        st.warning("SEU NÍVEL ATUAL É: OTIMIZAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE OTIMIZAÇÃO:** 
        Neste estágio, os processos estão plenamente integrados e gerenciados de maneira eficiente, com monitoramento contínuo e análise sistemática de desempenho. 
        A organização adota mecanismos formais de governança e controle, utilizando métricas para avaliação e aprimoramento das atividades. 
        A mitigação de riscos torna-se mais eficaz, com a implementação de políticas proativas para conformidade regulatória e excelência operacional. 
        O aprendizado organizacional é fomentado, garantindo a adaptação rápida a mudanças no ambiente interno e externo.
        """)
    elif total_porcentagem >= 91:
        nivel_atual = "EXCELÊNCIA"
        st.success("SEU NÍVEL ATUAL É: EXCELÊNCIA")
        st.info("""
        **NIVEL DE MATURIDADE EXCELÊNCIA:** 
        A organização alcança um nível de maturidade de referência, caracterizado por uma cultura de melhoria contínua e inovação. 
        Os processos são constantemente avaliados e aprimorados com base em análise de dados e benchmarking, garantindo máxima eficiência e alinhamento estratégico. 
        Há uma integração plena entre tecnologia, governança e gestão de riscos, promovendo uma operação resiliente e altamente adaptável às mudanças do mercado e do cenário regulatório. 
        O comprometimento com a excelência e a sustentabilidade impulsiona a organização a atuar como referência no setor.
        """)
    
    # Exibir a tabela de níveis de maturidade com o nível atual destacado
    exibir_tabela_niveis_maturidade(nivel_atual)

def mostrar_nivel_atual_por_grupo(grupo, valor_percentual):
    if valor_percentual < 26:
        nivel_atual = "INICIAL"
        st.warning(f"SEU NÍVEL ATUAL NO GRUPO '{grupo}' É: INICIAL")
        st.info("""
        **NIVEL DE MATURIDADE INICIAL:**
        Neste estágio, a organização opera de forma desestruturada, sem processos claramente definidos ou formalizados.
        As atividades são executadas de maneira reativa, sem padronização ou diretrizes estabelecidas, tornando a execução dependente do conhecimento tácito de indivíduos, em vez de uma abordagem institucionalizada.
        A ausência de controle efetivo e a inexistência de mecanismos de monitoramento resultam em vulnerabilidades operacionais e elevado risco de não conformidade regulatória.
        """)
    elif valor_percentual < 51:
        nivel_atual = "ORGANIZAÇÃO"
        st.warning(f"SEU NÍVEL ATUAL NO GRUPO '{grupo}' É: ORGANIZAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE ORGANIZAÇÃO:**
        A organização começa a estabelecer processos básicos, ainda que de maneira incipiente e pouco estruturada.
        Algumas diretrizes são documentadas e há um esforço para replicar práticas em diferentes áreas, embora a consistência na execução continue limitada.
        As atividades ainda dependem fortemente da experiência individual, e a governança sobre os processos é mínima, resultando em baixa previsibilidade e dificuldade na identificação e mitigação de riscos sistêmicos.
        """)
    elif valor_percentual < 71:
        nivel_atual = "CONSOLIDAÇÃO"
        st.warning(f"SEU NÍVEL ATUAL NO GRUPO '{grupo}' É: CONSOLIDAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE CONSOLIDAÇÃO:**
        A organização atinge um nível de maturidade em que os processos são formalmente documentados e seguidos de maneira estruturada.
        Existe uma clareza maior sobre as responsabilidades e papéis, o que reduz a dependência do conhecimento individual.
        A implementação de controles internos começa a ganhar robustez, permitindo um maior alinhamento com as diretrizes regulatórias e estratégicas.
        Indicadores de desempenho são introduzidos, permitindo um acompanhamento inicial da eficácia operacional, embora a cultura de melhoria contínua ainda esteja em desenvolvimento.
        """)
    elif valor_percentual < 90:
        nivel_atual = "OTIMIZAÇÃO"
        st.warning(f"SEU NÍVEL ATUAL NO GRUPO '{grupo}' É: OTIMIZAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE OTIMIZAÇÃO:**
        Neste estágio, os processos estão plenamente integrados e gerenciados de maneira eficiente, com monitoramento contínuo e análise sistemática de desempenho.
        A organização adota mecanismos formais de governança e controle, utilizando métricas para avaliação e aprimoramento das atividades.
        A mitigação de riscos torna-se mais eficaz, com a implementação de políticas proativas para conformidade regulatória e excelência operacional.
        O aprendizado organizacional é fomentado, garantindo a adaptação rápida a mudanças no ambiente interno e externo.
        """)
    elif valor_percentual >= 91:
        nivel_atual = "EXCELÊNCIA"
        st.success(f"SEU NÍVEL ATUAL NO GRUPO '{grupo}' É: EXCELÊNCIA")
        st.info("""
        **NIVEL DE MATURIDADE EXCELÊNCIA:**
        A organização alcança um nível de maturidade de referência, caracterizado por uma cultura de melhoria contínua e inovação.
        Os processos são constantemente avaliados e aprimorados com base em análise de dados e benchmarking, garantindo máxima eficiência e alinhamento estratégico.
        Há uma integração plena entre tecnologia, governança e gestão de riscos, promovendo uma operação resiliente e altamente adaptável às mudanças do mercado e do cenário regulatório.
        """)
    
    # Exibir a tabela de níveis de maturidade com o nível atual destacado
    exibir_tabela_niveis_maturidade(nivel_atual)

def validar_nivel_maturidade(soma_percentual, total_porcentagem):
    if soma_percentual < 26:
        st.warning("SEU NÍVEL ATUAL É: INICIAL")
        st.info("""
        **NIVEL DE MATURIDADE INICIAL:**
        Neste estágio, a organização opera de forma desestruturada, sem processos claramente definidos ou formalizados.
        As atividades são executadas de maneira reativa, sem padronização ou diretrizes estabelecidas, tornando a execução dependente do conhecimento tácito de indivíduos, em vez de uma abordagem institucionalizada.
        A ausência de controle efetivo e a inexistência de mecanismos de monitoramento resultam em vulnerabilidades operacionais e elevado risco de não conformidade regulatória.
        """)
    elif soma_percentual < 51:
        st.warning("SEU NÍVEL ATUAL É: ORGANIZAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE ORGANIZAÇÃO:**
        A organização começa a estabelecer processos básicos, ainda que de maneira incipiente e pouco estruturada.
        Algumas diretrizes são documentadas e há um esforço para replicar práticas em diferentes áreas, embora a consistência na execução continue limitada.
        As atividades ainda dependem fortemente da experiência individual, e a governança sobre os processos é mínima, resultando em baixa previsibilidade e dificuldade na identificação e mitigação de riscos sistêmicos.
        """)
    elif soma_percentual < 71:
        st.warning("SEU NÍVEL ATUAL É: CONSOLIDAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE CONSOLIDAÇÃO:**
        A organização atinge um nível de maturidade em que os processos são formalmente documentados e seguidos de maneira estruturada.
        Existe uma clareza maior sobre as responsabilidades e papéis, o que reduz a dependência do conhecimento individual.
        A implementação de controles internos começa a ganhar robustez, permitindo um maior alinhamento com as diretrizes regulatórias e estratégicas.
        Indicadores de desempenho são introduzidos, permitindo um acompanhamento inicial da eficácia operacional, embora a cultura de melhoria contínua ainda esteja em desenvolvimento.
        """)
    elif soma_percentual < 90:
        st.warning("SEU NÍVEL ATUAL É: OTIMIZAÇÃO")
        st.info("""
        **NIVEL DE MATURIDADE OTIMIZAÇÃO:**
        Neste estágio, os processos estão plenamente integrados e gerenciados de maneira eficiente, com monitoramento contínuo e análise sistemática de desempenho.
        A organização adota mecanismos formais de governança e controle, utilizando métricas para avaliação e aprimoramento das atividades.
        A mitigação de riscos torna-se mais eficaz, com a implementação de políticas proativas para conformidade regulatória e excelência operacional.
        O aprendizado organizacional é fomentado, garantindo a adaptação rápida a mudanças no ambiente interno e externo.
        """)
    elif soma_percentual >= 91:
        st.success("SEU NÍVEL ATUAL É: EXCELÊNCIA")
        st.info("""
        **NIVEL DE MATURIDADE EXCELÊNCIA:**
        A organização alcança um nível de maturidade de referência, caracterizado por uma cultura de melhoria contínua e inovação.
        Os processos são constantemente avaliados e aprimorados com base em análise de dados e benchmarking, garantindo máxima eficiência e alinhamento estratégico.
        Há uma integração plena entre tecnologia, governança e gestão de riscos, promovendo uma operação resiliente e altamente adaptável às mudanças do mercado e do cenário regulatório.
        """)

if "formulario_preenchido" not in st.session_state:
    st.session_state.formulario_preenchido = False
if "grupo_atual" not in st.session_state:
    st.session_state.grupo_atual = 0
if "respostas" not in st.session_state:
    st.session_state.respostas = {}
if "mostrar_graficos" not in st.session_state:
    st.session_state.mostrar_graficos = False

# Inicializar as variáveis fig_original e fig_normalizado para evitar erros
fig_original = None
fig_normalizado = None

if not st.session_state.formulario_preenchido:
    # Adicionando a imagem no início com tamanho reduzido
    col1, col2 = st.columns([1, 1])
    with col1:
        st.image("https://raw.githubusercontent.com/DaniloNs-creator/MATURITY/main/logo.png", width=300)
        st.header("DIAGNÓSTICO DE GESTÃO, GOVERNANÇA E CONTROLES")
        st.subheader("Preencha suas informações para iniciar:")

        nome = st.text_input("Nome")
        email = st.text_input("E-mail")
        empresa = st.text_input("Empresa")
        telefone = st.text_input("Telefone")
        if st.button("Prosseguir"):
            if nome and email and empresa and telefone:
                st.session_state.nome = nome
                st.session_state.email = email
                st.session_state.empresa = empresa
                st.session_state.telefone = telefone
                st.session_state.formulario_preenchido = True

                # Carregar respostas salvas, se existirem
                st.session_state.respostas = carregar_respostas(email)
                st.success("Informações preenchidas com sucesso! Você pode prosseguir para o questionário.")
            else:
                st.error("Por favor, preencha todos os campos antes de prosseguir.")

        # Bloco de apresentação profissional com background animado
        st.markdown("""
        <style>
        /* Fundo animado para o bloco de apresentação */
        .apresentacao-animada-bg {
            position: relative;
            overflow: hidden;
            background: linear-gradient(120deg, #f8fafc 60%, #e3e9f7 100%);
            border-radius: 18px;
            border: 1.5px solid #e0e0e0;
            padding: 32px 28px 22px 28px;
            margin-top: 18px;
            margin-bottom: 18px;
            box-shadow: 0 6px 24px rgba(44, 62, 80, 0.10);
            font-family: 'Segoe UI', 'Arial', sans-serif;
            z-index: 1;
        }
        /* Elementos animados no fundo */
        .apresentacao-animada-bg .bg-shape1,
        .apresentacao-animada-bg .bg-shape2,
        .apresentacao-animada-bg .bg-shape3 {
            position: absolute;
            border-radius: 50%;
            opacity: 0.18;
            z-index: 0;
            filter: blur(2px);
        }
        .apresentacao-animada-bg .bg-shape1 {
            width: 180px; height: 180px;
            background: #1976d2;
            top: -40px; left: -60px;
            animation: movebg1 8s infinite alternate;
        }
        .apresentacao-animada-bg .bg-shape2 {
            width: 120px; height: 120px;
            background: #43a047;
            bottom: -30px; right: -40px;
            animation: movebg2 10s infinite alternate;
        }
        .apresentacao-animada-bg .bg-shape3 {
            width: 90px; height: 90px;
            background: #fbc02d;
            top: 60px; right: 30px;
            animation: movebg3 12s infinite alternate;
        }
        @keyframes movebg1 {
            0% { transform: translateY(0) scale(1);}
            100% { transform: translateY(30px) scale(1.08);}
        }
        @keyframes movebg2 {
            0% { transform: translateX(0) scale(1);}
            100% { transform: translateX(-30px) scale(1.12);}
        }
        @keyframes movebg3 {
            0% { transform: translateY(0) translateX(0) scale(1);}
            100% { transform: translateY(-20px) translateX(20px) scale(1.05);}
        }
        .apresentacao-animada-bg h4 {
            color: #1a237e;
            margin-bottom: 14px;
            font-size: 1.25rem;
            font-weight: 700;
            z-index: 2;
            position: relative;
        }
        .apresentacao-animada-bg ul {
            margin-top: 0;
            margin-bottom: 0;
            padding-left: 18px;
            z-index: 2;
            position: relative;
        }
        .apresentacao-animada-bg li {
            margin-bottom: 6px;
            font-size: 1.05rem;
        }
        .apresentacao-animada-bg .dimensao {
            color: #0d47a1;
            font-weight: 600;
        }
        .apresentacao-animada-bg .subitem {
            color: #374151;
            font-size: 0.98rem;
        }
        .apresentacao-animada-bg p {
            margin-top: 16px;
            font-size: 1.08rem;
            color: #263238;
            z-index: 2;
            position: relative;
        }
        </style>
        <div class="apresentacao-animada-bg">
            <div class="bg-shape1"></div>
            <div class="bg-shape2"></div>
            <div class="bg-shape3"></div>
            <h4>Bem-vindo ao Diagnóstico de Maturidade Empresarial</h4>
            <p>
                Esta ferramenta foi desenvolvida para proporcionar uma avaliação estratégica do nível de maturidade da sua empresa em três dimensões essenciais:
            </p>
            <ul>
                <li class="dimensao">Gestão:
                    <ul>
                        <li class="subitem">Estrutura organizacional</li>
                        <li class="subitem">Eficiência financeira</li>
                    </ul>
                </li>
                <li class="dimensao">Governança:
                    <ul>
                        <li class="subitem">Gestão de processos</li>
                        <li class="subitem">Gestão de riscos</li>
                        <li class="subitem">Compliance regulatório</li>
                        <li class="subitem">Efetividade do canal de denúncias</li>
                    </ul>
                </li>
                <li class="dimensao">Áreas Operacionais:
                    <ul>
                        <li class="subitem">Recursos Humanos</li>
                        <li class="subitem">Tecnologia da Informação</li>
                        <li class="subitem">Gestão de compras e estoques</li>
                        <li class="subitem">Contabilidade e controles financeiros</li>
                        <li class="subitem">Logística e distribuição</li>
                    </ul>
                </li>
            </ul>
            <p>
                <b>Por que realizar este diagnóstico?</b><br>
                A análise integrada destes aspectos permite identificar pontos fortes, oportunidades de melhoria e priorizar ações para o crescimento sustentável do seu negócio. 
                Ao final, você receberá um relatório personalizado com recomendações práticas para elevar a maturidade da sua organização.
            </p>
            <p style="margin-top:10px; color:#1565c0;">
                <b>Confidencialidade garantida:</b> Todas as informações fornecidas serão tratadas com total sigilo e utilizadas exclusivamente para fins de diagnóstico e orientação estratégica.
            </p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.image("https://raw.githubusercontent.com/DaniloNs-creator/MATURITY/main/foto.jpg", use_container_width=True)
else:
    url_arquivo = "https://raw.githubusercontent.com/DaniloNs-creator/MATURITY/main/FOMULARIO.txt"
    try:
        response = requests.get(url_arquivo)
        response.raise_for_status()

        # Inicializar as variáveis para evitar erros
        categorias = []
        valores = []
        valores_normalizados = []
        lines = response.text.splitlines()
        data = []
        grupo_atual = None
        for line in lines:
            parts = line.strip().split(';')
            if len(parts) >= 2:
                classe = parts[0].strip()
                pergunta = parts[1].strip()

                if classe.isdigit():
                    grupo_atual = f"{classe} - {pergunta}"
                else:
                    if grupo_atual:
                        data.append({'grupo': grupo_atual, 'classe': classe, 'pergunta': pergunta})

        perguntas_df = pd.DataFrame(data)

        if perguntas_df.empty or not {'grupo', 'classe', 'pergunta'}.issubset(perguntas_df.columns):
            st.error("Certifique-se de que o arquivo TXT contém as colunas 'grupo', 'classe' e 'pergunta'.")
            st.write("Conteúdo do arquivo processado:", perguntas_df.head())
        else:
            perguntas_hierarquicas = {}
            for _, row in perguntas_df.iterrows():
                grupo = row['grupo']
                classe = str(row['classe'])
                pergunta = row['pergunta']

                if grupo not in perguntas_hierarquicas:
                    perguntas_hierarquicas[grupo] = {"titulo": grupo, "subitens": {}}

                perguntas_hierarquicas[grupo]["subitens"][classe] = pergunta

            grupos = list(perguntas_hierarquicas.keys())
            
            # Criando navegação por grupos
            with st.sidebar:
               
                # Corrigindo o caminho da imagem para o URL bruto do GitHub
                st.image("https://raw.githubusercontent.com/DaniloNs-creator/MATURITY/main/logo.png")
                st.title("Navegação por Grupos")
                
                tab1, tab2, tab3 = st.tabs([ "GESTÃO", "GOVERNANÇA", "SETORES"])
                
                
                
                with tab1:
                    
                    if st.button("**📊 Eficiência de Gestão**" if st.session_state.grupo_atual == 0 else "📊 Eficiência de Gestão"):
                        st.session_state.grupo_atual = 0
                    if st.button("**🏛️ Estruturas**" if st.session_state.grupo_atual == 1 else "🏛️ Estruturas"):
                        st.session_state.grupo_atual = 1    
                
                with tab2:
                    if st.button("**🔄 Gestão de Processos**" if st.session_state.grupo_atual == 2 else "🔄 Gestão de Processos"):
                        st.session_state.grupo_atual = 2
                    if st.button("**⚠️ Gestão de Riscos**" if st.session_state.grupo_atual == 3 else "⚠️ Gestão de Riscos"):
                        st.session_state.grupo_atual = 3
                    if st.button("**📝 Compliance**" if st.session_state.grupo_atual == 4 else "📝 Compliance"):
                        st.session_state.grupo_atual = 4
                    if st.button("**📢 Canal de Denúncias**" if st.session_state.grupo_atual == 5 else "📢 Canal de Denúncias"):
                        st.session_state.grupo_atual = 5
                    if st.button("**🏢 Governança Corporativa**" if st.session_state.grupo_atual == 6 else "🏢 Governança Corporativa"):
                        st.session_state.grupo_atual = 6
                
                with tab3:
                    if st.button("**👥 Recursos Humanos**" if st.session_state.grupo_atual == 7 else "👥 Recursos Humanos"):
                        st.session_state.grupo_atual = 7
                    if st.button("**💻 Tecnologia da Informação**" if st.session_state.grupo_atual == 8 else "💻 Tecnologia da Informação"):
                        st.session_state.grupo_atual = 8
                    if st.button("**🛒 Compras**" if st.session_state.grupo_atual == 9 else "🛒 Compras"):
                        st.session_state.grupo_atual = 9
                    if st.button("**📦 Estoques**" if st.session_state.grupo_atual == 10 else "📦 Estoques"):
                        st.session_state.grupo_atual = 10
                    if st.button("**💰 Contabilidade e Controle Financeiro**" if st.session_state.grupo_atual == 11 else "💰 Contabilidade e Controle Financeiro"):
                        st.session_state.grupo_atual = 11
                    if st.button("**🚚 Logística e Distribuição**" if st.session_state.grupo_atual == 12 else "🚚 Logística e Distribuição"):
                        st.session_state.grupo_atual = 12

                # Adicionar texto explicativo abaixo dos botões
                st.write("""
                Para garantir uma análise mais eficiente e resultados mais assertivos, recomendamos iniciar o diagnóstico pela aba 'Gestão', respondendo aos dois blocos de questões relacionados. 
                Em seguida, prossiga para 'Governança' e, por fim, 'Setores'. 

                No entanto, caso prefira, você pode navegar diretamente para qualquer aba específica de acordo com suas prioridades ou áreas de interesse imediato.
                """)

            grupo_atual = st.session_state.grupo_atual

            # Textos introdutórios para cada grupo
            TEXTO_GRUPO1 = """
            O preenchimento de uma Matriz de Maturidade de Gestão Financeira é essencial para avaliar a eficiência dos processos financeiros, identificar lacunas e estruturar um plano de melhoria contínua. Ela permite medir o nível de controle sobre orçamento, fluxo de caixa, investimentos e riscos, fornecendo uma visão clara da saúde financeira da empresa. Além disso, facilita a tomada de decisões estratégicas, ajudando a mitigar riscos, otimizar recursos e garantir a sustentabilidade do negócio a longo prazo. Empresas que utilizam essa matriz conseguem se adaptar melhor a mudanças e aprimorar sua competitividade.
            """
            TEXTO_GRUPO2 = """
            A avaliação da maturidade da estrutura de uma organização é um processo essencial para entender o nível de desenvolvimento e a eficácia das práticas de governança, gestão de riscos, compliance e processos organizacionais. Trata-se de um diagnóstico completo que permite identificar pontos fortes, fragilidades e oportunidades de melhoria em diferentes áreas estratégicas.
            """
            TEXTO_GRUPO3 = """
            O preenchimento desta seção permite avaliar a maturidade do programa de Compliance, garantindo que a organização esteja em conformidade com regulamentações e boas práticas éticas. Ajuda a prevenir riscos legais, fortalecer a cultura organizacional e demonstrar compromisso com a integridade corporativa.
            """
            TEXTO_GRUPO4 = """
            Responder a estas perguntas auxilia na identificação, monitoramento e mitigação de riscos que podem impactar a operação. Com uma gestão de riscos eficiente, a empresa minimiza perdas, melhora a tomada de decisão e se prepara para desafios internos e externos, garantindo maior resiliência operacional.
            """
            TEXTO_GRUPO5 = """
            Esta seção permite avaliar a eficiência e a padronização dos processos internos. Um bom gerenciamento de processos melhora a produtividade, reduz desperdícios e assegura entregas consistentes. Além disso, facilita a implementação de melhorias contínuas e a adaptação a novas exigências do mercado.
            """
            TEXTO_GRUPO6 = """
            A governança bem estruturada assegura transparência, ética e eficiência na gestão da empresa. Com este diagnóstico, é possível fortalecer a tomada de decisão, alinhar os interesses das partes interessadas e garantir um crescimento sustentável, reduzindo riscos e aumentando a confiança dos stakeholders.
            """
            TEXTO_GRUPO7 = """
            Esta seção mede a maturidade da gestão de pessoas, garantindo que a empresa valorize seus colaboradores e mantenha um ambiente produtivo e inclusivo. Um RH eficiente melhora a retenção de talentos, impulsiona a inovação e alinha os funcionários à cultura e estratégia organizacional.
            """
            TEXTO_GRUPO8 = """
            Responder a estas perguntas ajuda a avaliar o nível de digitalização e segurança da empresa. Uma TI bem estruturada melhora a eficiência operacional, protege dados sensíveis e impulsiona a inovação, garantindo que a organização esteja preparada para desafios tecnológicos e competitivos.
            """
            TEXTO_GRUPO9 = """
            Esta seção permite identificar boas práticas e oportunidades de melhoria na gestão financeira. Com um controle eficiente, a empresa assegura sustentabilidade, reduz riscos de inadimplência e fraudes, melhora a liquidez e otimiza investimentos, garantindo saúde financeira e crescimento sustentável.
            """
            TEXTO_GRUPO10 = """
            O diagnóstico nesta área assegura que as compras sejam estratégicas, alinhadas às necessidades da empresa e aos melhores preços e prazos. Com processos estruturados, a organização reduz custos, melhora a qualidade dos insumos e fortalece a relação com fornecedores confiáveis.
            """
            TEXTO_GRUPO11 = """
            Avaliar a gestão de estoques permite reduzir desperdícios, evitar faltas e garantir uma operação eficiente. Com controle adequado, a empresa melhora a previsibilidade, reduz custos de armazenagem e assegura disponibilidade de produtos, otimizando o fluxo operacional.
            """
            TEXTO_GRUPO12 = """
            Responder a estas perguntas possibilita otimizar a cadeia logística, garantindo entregas ágeis e redução de custos operacionais. Um bom planejamento melhora o nível de serviço, evita atrasos e assegura eficiência no transporte, impactando positivamente a satisfação do cliente.
            """
            TEXTO_GRUPO13 = """
            Esta seção avalia a transparência e conformidade da contabilidade empresarial. Um controle rigoroso das demonstrações financeiras assegura a correta apuração de resultados, garantindo confiança e credibilidade junto a investidores e órgãos reguladores.
            """

            # Lista de perguntas obrigatórias
            perguntas_obrigatorias = [
                "1.02", "1.06", "1.42", "1.03", "1.13", "1.14", "1.30", "1.12", "1.19", "1.25", "1.41", "1.43", "1.27", "1.35", "1.45", "1.20",
                "2.10", "2.01", "2.16", "2.23", "2.05", "2.08", "2.25", "2.29", "2.21", "2.22",
                "3.01", "3.04", "3.08", "3.11", "3.29", "3.38", "3.40", "3.42", "3.43",
                "4.01", "4.02", "4.03", "4.04", "4.05", "4.06", "4.07", "4.08", "4.09","4.10",
                "5.01", "5.03", "5.04", "5.07", "5.10", "5.32", "5.35", "5.40"
                "6.01", "6.02", "6.03", "6.04", "6.05", "6.06", "6.07", "6.08", "6.09","6.10", "6.11", "6.12",
                "7.01", "7.02", "7.03", "7.04", "7.05", "7.06", "7.07", "7.08", "7.09","7.10",
                "8.01", "8.02", "8.03", "8.04", "8.05", "8.06", "8.07", "8.08", "8.09","8.10","8.11","8.12","8.13","8.14","8.15","8.16","8.17",
                "9.01", "9.02", "9.03", "9.04", "9.05", "9.06", "9.07", "9.08", "9.09","9.10",
                "10.01", "10.02", "10.03", "10.04", "10.05", "10.06", "10.07", "10.08","10.09","10.10",
                "11.01", "11.02", "11.03", "11.04", "11.05", "11.06", "11.07", "11.08","11.09","11.10",
                "12.01", "12.02", "12.03", "12.04", "12.05", "12.06", "12.07", "12.08","12.09","12.10",
                "13.01", "13.02", "13.03", "13.04", "13.05", "13.06", "13.07", "13.08","13.09","13.10"
            ]

            # Grupos obrigatórios (4, 6, 7, 8, 9, 10, 11, 12, 13)
            grupos_obrigatorios = [
                "4 - Gestão de Riscos",
                "6 - Governança Corporativa",
                "7 - Recursos Humanos",
                "8 - Tecnologia da Informação",
                "9 - Compras",
                "10 - Estoques",
                "11 - Contabilidade e Controle Financeiro",
                "12 - Logística e Distribuição",
                "13 - Contabilidade e Controle Financeiro"
            ]

            if grupo_atual < len(grupos):
                grupo = grupos[grupo_atual]

                # Exibe o texto introdutório correspondente ao grupo atual
                if grupo.startswith("1 -"):
                    st.markdown(TEXTO_GRUPO1)
                elif grupo.startswith("2 -"):
                    st.markdown(TEXTO_GRUPO2)
                elif grupo.startswith("3 -"):
                    st.markdown(TEXTO_GRUPO3)
                elif grupo.startswith("4 -"):
                    st.markdown(TEXTO_GRUPO4)
                elif grupo.startswith("5 -"):
                    st.markdown(TEXTO_GRUPO5)
                elif grupo.startswith("6 -"):
                    st.markdown(TEXTO_GRUPO6)
                elif grupo.startswith("7 -"):
                    st.markdown(TEXTO_GRUPO7)
                elif grupo.startswith("8 -"):
                    st.markdown(TEXTO_GRUPO8)
                elif grupo.startswith("9 -"):
                    st.markdown(TEXTO_GRUPO9)
                elif grupo.startswith("10 -"):
                    st.markdown(TEXTO_GRUPO10)
                elif grupo.startswith("11 -"):
                    st.markdown(TEXTO_GRUPO11)
                elif grupo.startswith("12 -"):
                    st.markdown(TEXTO_GRUPO12)
                elif grupo.startswith("13 -"):
                    st.markdown(TEXTO_GRUPO13)

                st.write(f"### {perguntas_hierarquicas[grupo]['titulo']}")
                
                # Verifica se todas as perguntas obrigatórias foram respondidas
                todas_obrigatorias_preenchidas = True
                obrigatorias_no_grupo = []
                
                for subitem, subpergunta in perguntas_hierarquicas[grupo]["subitens"].items():
                    if subitem in perguntas_obrigatorias:
                        obrigatorias_no_grupo.append(subitem)
                        if st.session_state.respostas.get(subitem, "Selecione") == "Selecione":
                            todas_obrigatorias_preenchidas = False

                # Adicionando verificações para evitar erros ao acessar chaves inexistentes
                for subitem, subpergunta in perguntas_hierarquicas[grupo]["subitens"].items():
                    if subitem not in st.session_state.respostas:
                        st.session_state.respostas[subitem] = "Selecione"  # Inicializa com "Selecione"

                for subitem, subpergunta in perguntas_hierarquicas[grupo]["subitens"].items():
                    if subitem not in st.session_state.respostas:
                        st.session_state.respostas[subitem] = "Selecione"  # Inicializa com "Selecione"

                # Dividindo as perguntas em blocos de 10
                subitens = list(perguntas_hierarquicas[grupo]["subitens"].items())
                blocos = [subitens[i:i + 10] for i in range(0, len(subitens), 10)]

                for idx, bloco in enumerate(blocos):
                    # Verifica se todas as perguntas do bloco foram respondidas
                    bloco_preenchido = all(
                        st.session_state.respostas.get(subitem, "Selecione") != "Selecione"
                        for subitem, _ in bloco
                    )
                    # Destaca o bloco se estiver preenchido
                    bloco_titulo = f"Bloco {idx + 1} de perguntas"
                    if bloco_preenchido:
                        bloco_titulo = f"✅ **:green[{bloco_titulo}]**"
                    with st.expander(bloco_titulo, expanded=bloco_preenchido):
                        for subitem, subpergunta in bloco:
                            # Adiciona check se a pergunta foi respondida
                            respondida = st.session_state.respostas.get(subitem, "Selecione") != "Selecione"
                            check = " ✔️" if respondida else ""
                            if subitem in perguntas_obrigatorias:
                                pergunta_label = f"**:red[{subitem} - {subpergunta}]{check}** (OBRIGATÓRIO)"  # Destaca em vermelho
                            else:
                                pergunta_label = f"{subitem} - {subpergunta}{check}"

                            resposta = st.selectbox(
                                pergunta_label,
                                options=list(mapeamento_respostas.keys()),
                                index=list(mapeamento_respostas.keys()).index(st.session_state.respostas[subitem])
                            )
                            st.session_state.respostas[subitem] = resposta

                col1, col2, col3 = st.columns(3)
                with col1:
                    if st.button("⬅️ Voltar"):
                        if st.session_state.grupo_atual > 0:
                            st.session_state.grupo_atual -= 1
                            st.session_state.mostrar_graficos = False
                with col2:
                    if st.button("➡️ Prosseguir"):
                        # Verifica se todas as perguntas obrigatórias do grupo atual foram respondidas
                        obrigatorias_no_grupo = [
                            subitem for subitem in perguntas_hierarquicas[grupo]["subitens"].keys()
                            if subitem in perguntas_obrigatorias
                        ]
                        todas_obrigatorias_preenchidas = all(
                            st.session_state.respostas.get(subitem, "Selecione") != "Selecione"
                            for subitem in obrigatorias_no_grupo
                        )

                        if not todas_obrigatorias_preenchidas:
                            st.error(f"Ops...! Para concluir esse grupo você precisa revisar todas as perguntas obrigatórias: {', '.join(obrigatorias_no_grupo)}")
                        else:
                            # Avança para o próximo grupo
                            st.session_state.grupo_atual += 1
                            st.session_state.mostrar_graficos = False
                            st.success("Você avançou para o próximo grupo.")
                with col3:
                    if st.button("💾 Salvar Progresso"):
                        salvar_respostas(st.session_state.nome, st.session_state.email, st.session_state.respostas)
                    # Substituir os dois botões por um só
                    if st.button("📊 Gerar Gráficos e Enviar por Email"):
                        fig_original, fig_normalizado = gerar_graficos_radar(perguntas_hierarquicas, st.session_state.respostas)
                        if fig_original is None or fig_normalizado is None:
                            st.error("Os gráficos não foram gerados corretamente. Verifique os dados de entrada.")
                        else:
                            excel_data = exportar_questionario(st.session_state.respostas, perguntas_hierarquicas)
                            if enviar_email(st.session_state.email, excel_data, fig_original, fig_normalizado):
                                st.success(f"Relatório enviado com sucesso para o email {st.session_state.email}!")
                            st.session_state.mostrar_graficos = True

                if st.session_state.mostrar_graficos:
                    # Mensagem de Relatório de Progresso
                    grupo_atual_nome = grupos[st.session_state.grupo_atual]
                    respostas_numericas = {k: mapeamento_respostas[v] for k, v in st.session_state.respostas.items()}
                    soma_respostas = sum(respostas_numericas[subitem] for subitem in perguntas_hierarquicas[grupo_atual_nome]["subitens"].keys())
                    num_perguntas = len(perguntas_hierarquicas[grupo_atual_nome]["subitens"])
                    if num_perguntas > 0:
                        valor_percentual = (soma_respostas / (num_perguntas * 5)) * 100
                        nivel_atual = ""
                        if valor_percentual < 26:
                            nivel_atual = "INICIAL"
                        elif valor_percentual < 51:
                            nivel_atual = "ORGANIZAÇÃO"
                        elif valor_percentual < 71:
                            nivel_atual = "CONSOLIDAÇÃO"
                        elif valor_percentual < 90:
                            nivel_atual = "OTIMIZAÇÃO"
                        elif valor_percentual >= 91:
                            nivel_atual = "EXCELÊNCIA"

                        # Determinar os próximos blocos
                        proximos_blocos = grupos[st.session_state.grupo_atual + 1:] if st.session_state.grupo_atual + 1 < len(grupos) else []
                        proximos_blocos_texto = ", ".join(proximos_blocos) if proximos_blocos else "Nenhum bloco restante."

                        # Exibir a mensagem
                        st.markdown(f"""
                        ### Relatório de Progresso

                        Você completou o Bloco **{grupo_atual_nome}**. Os resultados indicam que o seu nível de maturidade neste bloco é classificado como: **{nivel_atual}**.

                        Para aprofundarmos a análise e oferecermos insights mais estratégicos, recomendamos que você complete também:

                        **{proximos_blocos_texto}**

                        Nossos consultores especializados receberão este relatório e entrarão em contato para agendar uma discussão personalizada. Juntos, identificaremos oportunidades de melhoria e traçaremos os próximos passos para otimizar os processos da sua organização.
                        """)

                    # Gerar gráficos
                    fig_original, fig_normalizado = gerar_graficos_radar(perguntas_hierarquicas, st.session_state.respostas)
                    if fig_original and fig_normalizado:
                        col1, col2 = st.columns(2)
                        with col1:
                            st.plotly_chart(fig_original, use_container_width=True)
                        with col2:
                            st.plotly_chart(fig_normalizado, use_container_width=True)

                        # Calcular e exibir o nível atual apenas para o grupo atual
                        mostrar_nivel_atual_por_grupo(grupo_atual_nome, valor_percentual)
            else:
                st.write("### Todas as perguntas foram respondidas!")
                if st.button("Gerar Gráfico Final"):
                    # Verifica se todas as perguntas obrigatórias foram respondidas
                    todas_obrigatorias_respondidas = True
                    obrigatorias_nao_respondidas = []
                    
                    for pergunta in perguntas_obrigatorias:
                        if pergunta not in st.session_state.respostas or st.session_state.respostas.get(pergunta, "Selecione") == "Selecione":
                            todas_obrigatorias_respondidas = False
                            obrigatorias_nao_respondidas.append(pergunta)
                    
                    # Verifica se todos os grupos obrigatórios foram completamente respondidos
                    grupos_obrigatorios_completos = True
                    grupos_incompletos = []
                    
                    for grupo_obrigatorio in grupos_obrigatorios:
                        if grupo_obrigatorio in perguntas_hierarquicas:
                            for subitem in perguntas_hierarquicas[grupo_obrigatorio]["subitens"].keys():
                                if subitem not in st.session_state.respostas or st.session_state.respostas.get(subitem, "Selecione") == "Selecione":
                                    grupos_obrigatorios_completos = False
                                    grupos_incompletos.append(grupo_obrigatorio)
                                    break
                    
                    if not todas_obrigatorias_respondidas or not grupos_obrigatorios_completos:
                        mensagem_erro = []
                        if not todas_obrigatorias_respondidas:
                            mensagem_erro.append(f"Perguntas obrigatórias não respondidas: {', '.join(obrigatorias_nao_respondidas)}")
                        if not grupos_obrigatorios_completos:
                            mensagem_erro.append(f"Grupos obrigatórios incompletos: {', '.join(set(grupos_incompletos))}")
                        st.error(" | ".join(mensagem_erro))
                    else:
                        # Adicionando logs para depuração
                        try:
                            respostas = {k: mapeamento_respostas.get(v, 0) for k, v in st.session_state.respostas.items()}
                            categorias = []
                            valores = []
                            valores_normalizados = []
                            soma_total_respostas = sum(respostas.values())
                            for item, conteudo in perguntas_hierarquicas.items():
                                soma_respostas = sum(respostas[subitem] for subitem in conteudo["subitens"].keys())
                                num_perguntas = len(conteudo["subitens"])
                                if num_perguntas > 0:
                                    valor_percentual = (soma_respostas / (num_perguntas * 5)) * 100
                                    valor_normalizado = (soma_respostas / valor_percentual) * 100 if valor_percentual > 0 else 0
                                    categorias.append(conteudo["titulo"])
                                    valores.append(valor_percentual)
                                    valores_normalizados.append(valor_normalizado)
                            if len(categorias) != len(valores) or len(categorias) != len(valores_normalizados):
                                st.error("Erro: As listas de categorias e valores têm tamanhos diferentes.")
                            else:
                                if categorias:
                                    valores_original = valores + valores[:1]
                                    categorias_original = categorias + categorias[:1]
                                    fig_original = go.Figure()
                                    fig_original.add_trace(go.Scatterpolar(
                                        r=valores_original,
                                        theta=categorias_original,
                                        fill='toself',
                                        name='Gráfico Original'
                                    ))
                                    fig_original.update_layout(
                                        polar=dict(
                                            radialaxis=dict(
                                                visible=True,
                                                range=[0, 100]
                                            )),
                                        showlegend=False
                                    )
                                    valores_normalizados_fechado = valores_normalizados + valores_normalizados[:1]
                                    fig_normalizado = go.Figure()
                                    fig_normalizado.add_trace(go.Scatterpolar(
                                        r=valores_normalizados_fechado,
                                        theta=categorias_original,
                                        fill='toself',
                                        name='Gráfico Normalizado'
                                    ))
                                    fig_normalizado.update_layout(
                                        polar=dict(
                                            radialaxis=dict(
                                                visible=True,
                                                range=[0, 100]
                                            )),
                                        showlegend=False
                                    )
                                    col1, col2 = st.columns(2)
                                    with col1:
                                        st.plotly_chart(fig_original, use_container_width=True)
                                        st.write("### Gráfico 1")
                                        df_grafico_original = pd.DataFrame({'Categoria': categorias, 'Porcentagem': valores})
                                        total_porcentagem = df_grafico_original['Porcentagem'].sum()
                                        df_grafico_original.loc['Total'] = ['Total', total_porcentagem]
                                        st.dataframe(df_grafico_original)

                                        if total_porcentagem < 26:
                                            st.warning("SEU NIVEL É INICIAL")
                                        elif total_porcentagem < 51:
                                            st.warning("SEU NIVEL É ORGANIZAÇÃO")
                                        elif total_porcentagem < 71:
                                            st.warning("SEU NIVEL É CONSOLIDAÇÃO")
                                        elif total_porcentagem < 90:
                                            st.warning("SEU NIVEL É OTIMIZAÇÃO")
                                        elif total_porcentagem >= 91:
                                            st.success("SEU NIVEL É EXCELÊNCIA")
                                    with col2:
                                        st.plotly_chart(fig_normalizado, use_container_width=True)
                                        st.write("### Gráfico 2")
                                        df_grafico_normalizado = pd.DataFrame({'Categoria': categorias, 'Porcentagem Normalizada': valores_normalizados})
                                        st.dataframe(df_grafico_normalizado)
                                    
                                    # Mostrar nível de maturidade completo
                                    mostrar_nivel_maturidade(total_porcentagem)
                                    
                                    excel_data = exportar_questionario(st.session_state.respostas, perguntas_hierarquicas)
                                    st.download_button(
                                        label="Exportar para Excel",
                                        data=excel_data,
                                        file_name="questionario_preenchido.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    )
                        except KeyError as e:
                            st.error(f"Erro ao acessar chave inexistente: {e}")
                            st.write("Estado atual das respostas:", st.session_state.respostas)
                            st.write("Perguntas obrigatórias:", perguntas_obrigatorias)
                            st.write("Perguntas hierárquicas:", perguntas_hierarquicas)
    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar o arquivo: {e}")

# Garantir que perguntas_hierarquicas esteja definido
if 'perguntas_hierarquicas' not in locals():
    perguntas_hierarquicas = {}

# Garantir que perguntas_obrigatorias esteja definido
if 'perguntas_obrigatorias' not in locals():
    perguntas_obrigatorias = []

# Garantir que todas as perguntas obrigatórias sejam inicializadas no dicionário de respostas
for grupo, conteudo in perguntas_hierarquicas.items():
    for subitem in conteudo["subitens"].keys():
        if subitem not in st.session_state.respostas:
            st.session_state.respostas[subitem] = "Selecione"  # Inicializa com "Selecione"

# Adicionando verificações para evitar erros ao acessar chaves inexistentes
try:
    respostas = {k: mapeamento_respostas.get(v, 0) for k, v in st.session_state.respostas.items()}
except KeyError as e:
    st.error(f"Erro ao acessar chave inexistente: {e}")
    st.write("Estado atual das respostas:", st.session_state.respostas)
    st.write("Perguntas hierárquicas:", perguntas_hierarquicas)
