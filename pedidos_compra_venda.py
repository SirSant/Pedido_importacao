import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import pyautogui
import threading
import time
import os
import re
import ctypes
from decimal import Decimal, InvalidOperation
from PIL import Image


# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

TEMPO_PREPARO = 5
DELAY_ENTRE_ITENS = 0.6
LINHA_INICIAL_COMPRA = 2
LINHA_INICIAL_VENDA = 6

# Ritmo de digitação: ajustado para dar tempo ao sistema de registrar cada campo.
INTERVALO_DIGITACAO = 0.05
PAUSA_APOS_DIGITAR = 0.30
PAUSA_ENTRE_TABS_CABECALHO = 0.20
PAUSA_APOS_GRUPO_TABS = 0.35

executando = False
arquivo_excel = None
produtos_carregados = []

pyautogui.PAUSE = 0.20
pyautogui.FAILSAFE = True

# ==========================================================
# CONFIGURAÇÃO DA PLANILHA DE COMPRA
# ==========================================================

ABAS_CAPACETES = ["SHARK", "KYT", "ORIGINE", "SUOMY"]
ABAS_PECAS = [
    "PARTS SHARK",
    "PARTS KYT",
    "PARTS ORIGINE",
    "PARTS SUOMY",
    "DUOTECH",
]

COL_CAP_CODIGO = 4
COL_CAP_DESCRICAO = 5
COL_TAMANHO_S = 6
COL_TAMANHO_M = 7
COL_TAMANHO_L = 8
COL_TAMANHO_XL = 9

COL_PEC_CODIGO = 2
COL_PEC_DESCRICAO = 3
COL_PEC_QUANTIDADE = 4

COL_DUO_CODIGO = 3
COL_DUO_DESCRICAO = 4
COL_DUO_QUANTIDADE = 5

# ==========================================================
# CONFIGURAÇÃO DA PLANILHA DE VENDA
# ==========================================================

ABA_VENDA = "Ajuste de Custo"
COL_VENDA_CODIGO = 6       # F
COL_VENDA_DESCRICAO = 7    # G
COL_VENDA_QUANTIDADE = 8   # H
COL_VENDA_CUSTO_TRADE = 11 # K

# ==========================================================
# INTERFACE / FUNÇÕES AUXILIARES
# ==========================================================

def registrar_log(mensagem):
    def atualizar():
        log_text.configure(state="normal")
        log_text.insert(tk.END, mensagem + "\n")
        log_text.see(tk.END)
        log_text.configure(state="disabled")

    if "janela" in globals() and janela.winfo_exists():
        janela.after(0, atualizar)


def normalizar(valor):
    if valor is None:
        return ""

    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)

    return str(valor).strip()


def formatar_numero(valor):
    if valor is None:
        return ""

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return str(valor).strip()


def formatar_decimal_br(valor, max_inteiros=None, max_decimais=2, campo="valor"):
    """Valida um número e devolve a representação com vírgula decimal."""

    if valor is None or str(valor).strip() == "":
        return ""

    texto = str(valor).strip().replace(",", ".")

    if not re.fullmatch(r"\d+(?:\.\d{1,2})?", texto):
        raise ValueError(
            f"{campo}: informe apenas números e, no máximo, "
            f"{max_decimais} casas decimais."
        )

    try:
        numero = Decimal(texto)
    except InvalidOperation:
        raise ValueError(f"{campo}: valor numérico inválido.")

    if numero < 0:
        raise ValueError(f"{campo}: não informe valores negativos.")

    parte_inteira, _, parte_decimal = format(numero, "f").partition(".")
    parte_decimal = parte_decimal.rstrip("0")

    if max_inteiros is not None and len(parte_inteira) > max_inteiros:
        raise ValueError(
            f"{campo}: máximo de {max_inteiros} dígitos antes da vírgula."
        )

    if len(parte_decimal) > max_decimais:
        raise ValueError(
            f"{campo}: máximo de {max_decimais} casas decimais."
        )

    return parte_inteira + ("," + parte_decimal if parte_decimal else "")


def preparar_codigo_cabecalho(valor, nome, tamanho):
    texto = str(valor).strip()

    if not texto or not texto.isdigit():
        raise ValueError(f"{nome}: informe somente números.")

    if len(texto) > tamanho:
        raise ValueError(f"{nome}: informe no máximo {tamanho} dígitos.")

    return texto.zfill(tamanho)


def preparar_operacao(valor):
    texto = str(valor).strip()

    if not texto or not texto.isdigit():
        raise ValueError("Operação: informe um código numérico.")

    if len(texto) > 2:
        raise ValueError("Operação: informe no máximo 2 dígitos.")

    return texto.zfill(2)


def colar_texto_unicode(texto):
    """Usa a área de transferência do Tk para colar textos com acentos."""

    concluido = threading.Event()
    erros = []

    def preparar_clipboard():
        try:
            janela.clipboard_clear()
            janela.clipboard_append(texto)
            janela.update()

        except Exception as erro:
            erros.append(erro)

        finally:
            concluido.set()

    janela.after(0, preparar_clipboard)

    if not concluido.wait(timeout=8):
        raise RuntimeError(
            "Não foi possível preparar a mensagem para colagem."
        )

    if erros:
        raise RuntimeError(
            f"Falha ao preparar a mensagem: {erros[0]}"
        )

    # Dá tempo para o sistema focar o campo antes de colar e processar o texto.
    time.sleep(0.35)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.55)


def selecionar_arquivo():
    global arquivo_excel

    caminho = filedialog.askopenfilename(
        title="Selecionar planilha de pedido",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if caminho:
        arquivo_excel = caminho
        nome = os.path.basename(caminho)

        lbl_arquivo.config(
            text=nome,
            foreground="#dce6f2"
        )

        registrar_log(f"Planilha selecionada: {caminho}")


# ==========================================================
# LEITURA DE PEDIDOS DE COMPRA
# ==========================================================

def gerar_produtos_capacetes(codigo, descricao, quantidades):
    tamanhos = ["S", "M", "L", "XL"]
    produtos = []

    if "." not in codigo:
        raise ValueError(
            f"O código de capacete '{codigo}' não possui sufixo separado por ponto."
        )

    codigo_base, sufixo = codigo.rsplit(".", 1)

    for tamanho, quantidade in zip(tamanhos, quantidades):

        if quantidade in (None, 0, ""):
            continue

        produtos.append({
            "codigo": f"{codigo_base}{tamanho}.{sufixo}",
            "descricao": normalizar(descricao),
            "quantidade": quantidade,
            "preco_unit": None,
            "categoria": "CAPACETE",
        })

    return produtos


def ler_capacetes(planilha, nome_aba):
    produtos = []

    registrar_log(f"Lendo aba: {nome_aba}")

    for linha in range(
        LINHA_INICIAL_COMPRA,
        planilha.max_row + 1
    ):

        codigo = planilha.cell(
            linha, COL_CAP_CODIGO
        ).value

        if codigo in (None, ""):
            continue

        descricao = planilha.cell(
            linha, COL_CAP_DESCRICAO
        ).value

        quantidades = [
            planilha.cell(linha, COL_TAMANHO_S).value,
            planilha.cell(linha, COL_TAMANHO_M).value,
            planilha.cell(linha, COL_TAMANHO_L).value,
            planilha.cell(linha, COL_TAMANHO_XL).value,
        ]

        produtos.extend(
            gerar_produtos_capacetes(
                normalizar(codigo),
                descricao,
                quantidades
            )
        )

    return produtos


def ler_pecas(planilha, nome_aba):
    produtos = []

    registrar_log(f"Lendo aba: {nome_aba}")

    for linha in range(
        LINHA_INICIAL_COMPRA,
        planilha.max_row + 1
    ):

        codigo = planilha.cell(
            linha, COL_PEC_CODIGO
        ).value

        descricao = planilha.cell(
            linha, COL_PEC_DESCRICAO
        ).value

        quantidade = planilha.cell(
            linha, COL_PEC_QUANTIDADE
        ).value

        if codigo in (None, "") or quantidade in (None, 0, ""):
            continue

        produtos.append({
            "codigo": normalizar(codigo),
            "descricao": normalizar(descricao),
            "quantidade": quantidade,
            "preco_unit": None,
            "categoria": "PEÇA",
        })

    return produtos


def ler_duotech(planilha, nome_aba):
    produtos = []

    registrar_log(f"Lendo aba: {nome_aba}")

    for linha in range(
        LINHA_INICIAL_COMPRA,
        planilha.max_row + 1
    ):

        codigo = planilha.cell(
            linha, COL_DUO_CODIGO
        ).value

        descricao = planilha.cell(
            linha, COL_DUO_DESCRICAO
        ).value

        quantidade = planilha.cell(
            linha, COL_DUO_QUANTIDADE
        ).value

        if codigo in (None, "") or quantidade in (None, 0, ""):
            continue

        produtos.append({
            "codigo": normalizar(codigo),
            "descricao": normalizar(descricao),
            "quantidade": quantidade,
            "preco_unit": None,
            "categoria": "PEÇA",
        })

    return produtos


def montar_lista_compra(workbook):
    produtos = []

    for nome_aba in ABAS_CAPACETES:

        if nome_aba in workbook.sheetnames:
            produtos.extend(
                ler_capacetes(workbook[nome_aba], nome_aba)
            )

    for nome_aba in ABAS_PECAS:

        if nome_aba not in workbook.sheetnames:
            continue

        if nome_aba == "DUOTECH":
            produtos.extend(
                ler_duotech(workbook[nome_aba], nome_aba)
            )

        else:
            produtos.extend(
                ler_pecas(workbook[nome_aba], nome_aba)
            )

    return produtos


# ==========================================================
# LEITURA DE PEDIDOS DE VENDA
# ==========================================================

def montar_lista_venda(workbook):

    if ABA_VENDA not in workbook.sheetnames:
        raise ValueError(
            f"A planilha não possui a aba '{ABA_VENDA}'."
        )

    planilha = workbook[ABA_VENDA]
    produtos = []

    registrar_log(f"Lendo aba: {ABA_VENDA}")

    for linha in range(
        LINHA_INICIAL_VENDA,
        planilha.max_row + 1
    ):

        codigo = planilha.cell(
            linha, COL_VENDA_CODIGO
        ).value

        descricao = planilha.cell(
            linha, COL_VENDA_DESCRICAO
        ).value

        quantidade = planilha.cell(
            linha, COL_VENDA_QUANTIDADE
        ).value

        custo_trade = planilha.cell(
            linha, COL_VENDA_CUSTO_TRADE
        ).value

        if codigo in (None, "") or quantidade in (None, 0, ""):
            continue

        produtos.append({
            "codigo": normalizar(codigo),
            "descricao": normalizar(descricao),
            "quantidade": quantidade,
            "preco_unit": custo_trade,
            "categoria": "VENDA",
        })

    return produtos


def montar_lista_produtos(workbook, tipo_pedido):

    if tipo_pedido == "Compra":
        return montar_lista_compra(workbook)

    return montar_lista_venda(workbook)


# ==========================================================
# DIGITAÇÃO NO SISTEMA
# ==========================================================

def digitar(texto):
    """Digita com intervalo entre caracteres e aguarda o sistema processar o campo."""

    if texto is None or texto == "":
        return

    pyautogui.write(
        str(texto),
        interval=INTERVALO_DIGITACAO
    )

    time.sleep(PAUSA_APOS_DIGITAR)


def avancar_tabs_cabecalho(quantidade=1):
    """Avança campo a campo, evitando enviar vários TABs em sequência rápida."""

    for _ in range(quantidade):

        if not executando:
            return

        pyautogui.press("tab")
        time.sleep(PAUSA_ENTRE_TABS_CABECALHO)

    time.sleep(PAUSA_APOS_GRUPO_TABS)


def digitar_pedido(produtos, zerar_lista, digitar_preco, operacao):

    for produto in produtos:

        if not executando:
            break

        codigo = produto["codigo"]
        quantidade = formatar_numero(produto["quantidade"])
        preco_unit = produto.get("preco_unit")

        registrar_log(
            f"Digitando: {codigo} | Qtd: {quantidade} | "
            f"Custo unit.: {formatar_numero(preco_unit) or 'não informado'} | "
            f"Operação: {operacao}"
        )

        # Código do produto
        digitar(codigo)

        # O sistema avança sozinho quando o código tem 15 caracteres.
        if len(codigo) != 15:
            pyautogui.press("tab")

        for _ in range(3):
            pyautogui.press("right")

        # Quantidade
        digitar(quantidade)
        pyautogui.press("tab")

        # Preço de lista
        if zerar_lista:
            digitar("0")

        pyautogui.press("tab")

        # Preço unitário
        if digitar_preco and preco_unit not in (None, ""):
            digitar(formatar_numero(preco_unit))

        elif digitar_preco and preco_unit in (None, ""):
            registrar_log(
                f"AVISO: {codigo} não possui custo unitário na planilha; "
                "campo não preenchido."
            )

        pyautogui.press("tab")

        for _ in range(2):
            pyautogui.press("right")

        # Operação escolhida na interface
        digitar(operacao)
        pyautogui.press("down")

        time.sleep(DELAY_ENTRE_ITENS)

def verificar_aviso_cliente():
    """Identifica o aviso de crédito e confirma automaticamente."""

    caminho_imagem = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "aviso_cliente.png"
    )

    try:
        # Abre a imagem usando Pillow, que aceita caminhos com caracteres especiais.
        with Image.open(caminho_imagem) as arquivo:
            imagem_referencia = arquivo.convert("RGB")

    except Exception as erro:
        raise RuntimeError(
            f"Não foi possível abrir aviso_cliente.png: {erro}"
        )

    limite = time.monotonic() + 2

    while executando and time.monotonic() < limite:

        aviso = pyautogui.locateOnScreen(
            imagem_referencia,
            confidence=0.85
        )

        if aviso is not None:
            registrar_log(
                "AVISO DE CRÉDITO IDENTIFICADO. Confirmando com Enter."
            )

            pyautogui.press("enter")
            time.sleep(0.3)
            return True

        time.sleep(0.1)

    registrar_log("Nenhum aviso de crédito identificado.")
    return False


# ==========================================================
# DIGITAÇÃO DE CABEÇALHO E ITENS DE VENDA
# ==========================================================

def digitar_cabecalho_venda(cabecalho):
    """Preenche o cabeçalho com pausas entre digitação e navegação dos campos."""

    registrar_log(
        "Preenchendo cabeçalho do pedido de venda..."
    )

    # Cliente
    digitar(cabecalho["cliente"])
    avancar_tabs_cabecalho(1)
    verificar_aviso_cliente()
    avancar_tabs_cabecalho(2)

    # Peso bruto
    digitar(cabecalho["peso_bruto"])
    avancar_tabs_cabecalho(1)

    # Volume 1
    digitar(cabecalho["volume1"])
    avancar_tabs_cabecalho(2)

    # Transportadora
    digitar(cabecalho["transportadora"])
    avancar_tabs_cabecalho(3)

    # Condição de pagamento e forma de pagamento fixa em X.
    digitar(cabecalho["cond_pag"])
    digitar("X")

    registrar_log(
        "Navegando até o campo da mensagem da nota..."
    )

    avancar_tabs_cabecalho(27)
    time.sleep(0.50)

    # Mensagem da nota: aguarda o campo estabilizar antes e depois da colagem.
    if cabecalho["mensagem"]:
        colar_texto_unicode(cabecalho["mensagem"])

    avancar_tabs_cabecalho(1)

    # Código fixo indicado no roteiro.
    digitar("008")
    avancar_tabs_cabecalho(1)

    time.sleep(0.50)

    registrar_log(
        "Cabeçalho preenchido. Clique na primeira linha de produtos no sistema para continuar."
    )

    aguardar_clique_primeira_linha()


def aguardar_clique_primeira_linha():
    """Espera o usuário clicar na primeira linha do grid, sem inventar coordenadas."""

    pyautogui.click(x=120, y=640)
    time.sleep(0.1)

    registrar_log(
        "AGUARDANDO: Selecionando a primeira linha do pedido"
    )

def digitar_pedido_venda(produtos, operacao):
    """Executa o ciclo de item de venda descrito na planilha de exemplo."""

    for indice, produto in enumerate(produtos, start=1):

        if not executando:
            break

        codigo = normalizar(produto["codigo"])
        quantidade = formatar_numero(produto["quantidade"])

        if not quantidade.isdigit() or len(quantidade) > 6:
            raise ValueError(
                f"Quantidade inválida para o produto {codigo}: "
                f"{quantidade!r}. Máximo de 6 dígitos."
            )

        if len(codigo) > 15:
            raise ValueError(
                f"Código {codigo} possui mais de 15 caracteres."
            )

        preco_unit = produto.get("preco_unit")

        preco_digitado = formatar_decimal_br(
            preco_unit,
            max_inteiros=4,
            max_decimais=2,
            campo=f"Custo Trade do produto {codigo}"
        ) if preco_unit not in (None, "") else ""

        registrar_log(
            f"Venda {indice}/{len(produtos)} | {codigo} | "
            f"Qtd: {quantidade} | "
            f"Custo Trade (K): {preco_digitado or 'não informado'} | "
            f"Operação: {operacao}"
        )

        # Código (coluna F): 15 caracteres avançam automaticamente; abaixo disso, TAB manual.
        digitar(codigo)

        if len(codigo) != 15:
            pyautogui.press("tab")

        # Três setas à direita até a quantidade (coluna H).
        pyautogui.press("right", presses=3)
        digitar(quantidade)

        # Preço de lista fixo em zero.
        pyautogui.press("tab")
        digitar("0")
        pyautogui.press("tab")

        # Preço unitário lido da coluna K (Custo Trade).
        if preco_digitado:
            digitar(preco_digitado)

        else:
            registrar_log(
                f"AVISO: {codigo} está sem Custo Trade na coluna K; "
                "campo deixado em branco."
            )

        pyautogui.press("tab")

        # Operação e campos seguintes conforme o algoritmo da planilha.
        pyautogui.press("right", presses=2)
        digitar(operacao)
        pyautogui.press("right", presses=2)
        digitar("0")
        pyautogui.press("down")
        pyautogui.press("right")

        time.sleep(DELAY_ENTRE_ITENS)


# ==========================================================
# EXECUÇÃO
# ==========================================================

def executar_automacao(
    tipo_pedido,
    zerar_lista,
    digitar_preco,
    operacao,
    cabecalho=None
):
    global executando

    workbook = None

    try:
        workbook = load_workbook(
            arquivo_excel,
            data_only=True
        )

        registrar_log("=" * 48)
        registrar_log(
            f"TIPO DE PEDIDO: {tipo_pedido.upper()}"
        )

        registrar_log(
            "Lendo e validando a planilha..."
        )

        produtos = montar_lista_produtos(
            workbook,
            tipo_pedido
        )

        if not produtos:
            raise ValueError(
                "Nenhum produto válido foi encontrado na planilha."
            )

        if tipo_pedido == "Compra" and digitar_preco:
            registrar_log(
                "ATENÇÃO: a planilha de compra atual não possui "
                "coluna de preço unitário configurada. "
                "Os preços desses itens serão pulados."
            )

        registrar_log("-" * 48)
        registrar_log(
            f"Produtos identificados: {len(produtos)}"
        )

        for produto in produtos:
            registrar_log(
                f"{produto['codigo']} | "
                f"{produto['quantidade']} | "
                f"{produto['categoria']}"
            )

        registrar_log("-" * 48)

        registrar_log(
            f"Iniciando em {TEMPO_PREPARO} segundos..."
        )

        for segundo in range(TEMPO_PREPARO, 0, -1):

            if not executando:
                return

            registrar_log(f"{segundo}...")
            time.sleep(1)

        if not executando:
            return

        if tipo_pedido == "Venda":

            digitar_cabecalho_venda(
                cabecalho or {}
            )

            if executando:
                digitar_pedido_venda(
                    produtos,
                    operacao
                )

        else:
            # Mantém o fluxo de compra original.
            digitar_pedido(
                produtos,
                zerar_lista,
                digitar_preco,
                operacao
            )

        if executando:
            registrar_log(
                "Processamento finalizado."
            )

        else:
            registrar_log(
                "Processamento interrompido pelo usuário."
            )

    except pyautogui.FailSafeException:
        registrar_log(
            "AUTOMAÇÃO INTERROMPIDA: mouse levado ao canto superior esquerdo."
        )

        janela.after(
            0,
            lambda: messagebox.showwarning(
                "Automação interrompida",
                "O PyAutoGUI foi interrompido pelo recurso de segurança."
            )
        )

    except Exception as erro:
        mensagem = str(erro)

        registrar_log(
            f"ERRO: {mensagem}"
        )

        janela.after(
            0,
            lambda msg=mensagem: messagebox.showerror(
                "Erro",
                msg
            )
        )

    finally:
        if workbook:
            workbook.close()

        executando = False

        janela.after(
            0,
            atualizar_estado_botoes
        )


def atualizar_estado_botoes():
    botao_iniciar.config(
        state="normal"
    )

    botao_parar.config(
        state="disabled"
    )


def coletar_dados_cabecalho_venda():

    cliente = preparar_codigo_cabecalho(
        campo_cliente.get(),
        "Cliente",
        6
    )

    transportadora = preparar_codigo_cabecalho(
        campo_transportadora.get(),
        "Transportadora",
        6
    )

    cond_pag = preparar_codigo_cabecalho(
        campo_cond_pag.get(),
        "Cond. Pag",
        3
    )

    volume1 = preparar_codigo_cabecalho(
        campo_volume1.get(),
        "Volume 1",
        5
    )

    peso_bruto = formatar_decimal_br(
        campo_peso_bruto.get(),
        max_inteiros=None,
        max_decimais=2,
        campo="Peso bruto"
    )

    if not peso_bruto:
        raise ValueError(
            "Peso bruto: preencha o campo."
        )

    return {
        "cliente": cliente,
        "transportadora": transportadora,
        "cond_pag": cond_pag,
        "peso_bruto": peso_bruto,
        "volume1": volume1,
        "mensagem": campo_mensagem.get().strip(),
    }


def iniciar_automacao():
    global executando

    if executando:
        return

    if not arquivo_excel:
        messagebox.showwarning(
            "Atenção",
            "Selecione uma planilha primeiro."
        )
        return

    tipo = tipo_pedido.get()

    if tipo not in ("Compra", "Venda"):
        messagebox.showwarning(
            "Atenção",
            "Selecione Compra ou Venda."
        )
        return

    try:
        operacao = preparar_operacao(
            campo_operacao.get()
        )

        cabecalho = (
            coletar_dados_cabecalho_venda()
            if tipo == "Venda"
            else None
        )

    except ValueError as erro:
        messagebox.showwarning(
            "Dados inválidos",
            str(erro)
        )
        return

    zerar_lista = bool(
        zerar_preco_lista.get()
    )

    preencher_preco = bool(
        digitar_preco_unit.get()
    )

    if tipo == "Compra" and preencher_preco:

        confirmar = messagebox.askyesno(
            "Preço unitário",
            "A leitura atual da planilha de compra não possui "
            "coluna de preço unitário configurada. "
            "Os campos de preço unitário ficarão sem preenchimento. "
            "Deseja continuar?"
        )

        if not confirmar:
            return

    if tipo == "Venda":

        confirmar = messagebox.askyesno(
            "Confirme o cabeçalho da venda",
            "O programa preencherá os dados do cabeçalho, "
            "colará a mensagem e digitará o código 008. "
            "Em seguida aguardará um clique na primeira linha "
            "de produtos.\n\nDeseja continuar?"
        )

        if not confirmar:
            return

    executando = True

    botao_iniciar.config(
        state="disabled"
    )

    botao_parar.config(
        state="normal"
    )

    threading.Thread(
        target=executar_automacao,
        args=(
            tipo,
            zerar_lista,
            preencher_preco,
            operacao,
            cabecalho
        ),
        daemon=True
    ).start()


def parar_automacao():
    global executando

    if executando:
        executando = False

        registrar_log(
            "Solicitando interrupção. "
            "O processo será encerrado após a etapa atual."
        )


def limpar_log():
    log_text.configure(state="normal")
    log_text.delete("1.0", tk.END)
    log_text.configure(state="disabled")


# ==========================================================
# INTERFACE MODERNA - VISUAL ORIGINAL PRESERVADO
# ==========================================================

janela = tk.Tk()
janela.title(
    "Winner Motors | Automação de Pedidos"
)

janela.geometry("980x740")
janela.minsize(900, 680)
janela.configure(bg="#111827")

style = ttk.Style()
style.theme_use("clam")

BG = "#111827"
PANEL = "#1f2937"
PANEL_ALT = "#263244"
TEXT = "#f9fafb"
MUTED = "#9ca3af"
ACCENT = "#0d9488"
ACCENT_HOVER = "#0f766e"
DANGER = "#dc2626"
BORDER = "#374151"

style.configure(
    "TFrame",
    background=BG
)

style.configure(
    "Panel.TFrame",
    background=PANEL
)

style.configure(
    "TLabel",
    background=BG,
    foreground=TEXT,
    font=("Segoe UI", 10)
)

style.configure(
    "Title.TLabel",
    background=BG,
    foreground=TEXT,
    font=("Segoe UI", 20, "bold")
)

style.configure(
    "Subtitle.TLabel",
    background=BG,
    foreground=MUTED,
    font=("Segoe UI", 9)
)

style.configure(
    "PanelTitle.TLabel",
    background=PANEL,
    foreground=TEXT,
    font=("Segoe UI", 11, "bold")
)

style.configure(
    "Muted.TLabel",
    background=PANEL,
    foreground=MUTED,
    font=("Segoe UI", 9)
)

style.configure(
    "TCombobox",
    fieldbackground=PANEL_ALT,
    background=PANEL_ALT,
    foreground=TEXT,
    arrowcolor=TEXT,
    bordercolor=BORDER,
    padding=7
)

style.map(
    "TCombobox",
    fieldbackground=[
        ("readonly", PANEL_ALT)
    ],
    foreground=[
        ("readonly", TEXT)
    ]
)

style.configure(
    "Accent.TButton",
    background=ACCENT,
    foreground="white",
    font=("Segoe UI", 10, "bold"),
    padding=(16, 10),
    borderwidth=0
)

style.map(
    "Accent.TButton",
    background=[
        ("active", ACCENT_HOVER),
        ("disabled", "#4b5563")
    ]
)

style.configure(
    "Danger.TButton",
    background=DANGER,
    foreground="white",
    font=("Segoe UI", 10, "bold"),
    padding=(16, 10),
    borderwidth=0
)

style.map(
    "Danger.TButton",
    background=[
        ("active", "#b91c1c"),
        ("disabled", "#4b5563")
    ]
)

style.configure(
    "Secondary.TButton",
    background=PANEL_ALT,
    foreground=TEXT,
    font=("Segoe UI", 9),
    padding=(12, 8),
    borderwidth=0
)

style.map(
    "Secondary.TButton",
    background=[
        ("active", BORDER)
    ]
)

# Variáveis dos controles

tipo_pedido = tk.StringVar(value="Compra")
zerar_preco_lista = tk.BooleanVar(value=False)
digitar_preco_unit = tk.BooleanVar(value=False)

container = ttk.Frame(
    janela,
    padding=20
)

container.pack(
    fill="both",
    expand=True
)

# Cabeçalho

cabecalho = ttk.Frame(container)

cabecalho.pack(
    fill="x",
    pady=(0, 12)
)

ttk.Label(
    cabecalho,
    text="Automação de Pedidos",
    style="Title.TLabel"
).pack(anchor="w")

ttk.Label(
    cabecalho,
    text="Winner Motors  •  Cadastro automatizado de pedidos de compra e venda",
    style="Subtitle.TLabel"
).pack(
    anchor="w",
    pady=(4, 0)
)

# Configuração do tipo e da operação + opções de preço

linha_config = ttk.Frame(container)

linha_config.pack(
    fill="x",
    pady=(0, 10)
)

linha_config.columnconfigure(0, weight=1)
linha_config.columnconfigure(1, weight=1)

painel_pedido = tk.Frame(
    linha_config,
    bg=PANEL,
    highlightthickness=1,
    highlightbackground=BORDER
)

painel_pedido.grid(
    row=0,
    column=0,
    sticky="nsew",
    padx=(0, 6)
)

tk.Label(
    painel_pedido,
    text="CONFIGURAÇÃO DO PEDIDO",
    bg=PANEL,
    fg=TEXT,
    font=("Segoe UI", 10, "bold")
).pack(
    anchor="w",
    padx=16,
    pady=(12, 8)
)

linha_tipo_operacao = tk.Frame(
    painel_pedido,
    bg=PANEL
)

linha_tipo_operacao.pack(
    fill="x",
    padx=16,
    pady=(0, 14)
)

col_tipo = tk.Frame(
    linha_tipo_operacao,
    bg=PANEL
)

col_tipo.pack(
    side="left",
    fill="x",
    expand=True,
    padx=(0, 12)
)

tk.Label(
    col_tipo,
    text="Tipo de pedido",
    bg=PANEL,
    fg=MUTED,
    font=("Segoe UI", 9)
).pack(anchor="w")

combo_tipo = ttk.Combobox(
    col_tipo,
    textvariable=tipo_pedido,
    values=["Compra", "Venda"],
    state="readonly",
    font=("Segoe UI", 10),
    width=16
)

combo_tipo.pack(
    fill="x",
    pady=(5, 0)
)

col_operacao = tk.Frame(
    linha_tipo_operacao,
    bg=PANEL
)

col_operacao.pack(
    side="left",
    fill="x",
    expand=True
)

tk.Label(
    col_operacao,
    text="Operação (numérica)",
    bg=PANEL,
    fg=MUTED,
    font=("Segoe UI", 9)
).pack(anchor="w")

campo_operacao = tk.Entry(
    col_operacao,
    width=12,
    justify="center",
    bg=PANEL_ALT,
    fg=TEXT,
    insertbackground=TEXT,
    relief="flat",
    font=("Segoe UI", 10),
    highlightthickness=1,
    highlightbackground=BORDER,
    highlightcolor=ACCENT
)

campo_operacao.insert(0, "01")

campo_operacao.pack(
    fill="x",
    ipady=7,
    pady=(5, 0)
)

painel_preco = tk.Frame(
    linha_config,
    bg=PANEL,
    highlightthickness=1,
    highlightbackground=BORDER
)

painel_preco.grid(
    row=0,
    column=1,
    sticky="nsew",
    padx=(6, 0)
)

tk.Label(
    painel_preco,
    text="OPÇÕES DE PREÇO (COMPRAS)",
    bg=PANEL,
    fg=TEXT,
    font=("Segoe UI", 10, "bold")
).pack(
    anchor="w",
    padx=16,
    pady=(12, 8)
)


def criar_check(parent, texto, variavel):

    return tk.Checkbutton(
        parent,
        text=texto,
        variable=variavel,
        bg=PANEL,
        fg=TEXT,
        activebackground=PANEL,
        activeforeground=TEXT,
        selectcolor=PANEL_ALT,
        font=("Segoe UI", 9),
        anchor="w",
        relief="flat",
        highlightthickness=0
    )


check_zerar_lista = criar_check(
    painel_preco,
    "Zerar preço de lista",
    zerar_preco_lista
)

check_zerar_lista.pack(
    fill="x",
    padx=12,
    pady=(0, 4)
)

check_preco_unit = criar_check(
    painel_preco,
    "Digitar preço unitário",
    digitar_preco_unit
)

check_preco_unit.pack(
    fill="x",
    padx=12,
    pady=(0, 4)
)

label_preco = tk.Label(
    painel_preco,
    text="Na venda, o preço de lista será 0 e o preço unitário virá da coluna K.",
    bg=PANEL,
    fg=MUTED,
    font=("Segoe UI", 8),
    wraplength=360,
    justify="left"
)

label_preco.pack(
    anchor="w",
    padx=16,
    pady=(2, 10)
)

# Dados de cabeçalho - usados no fluxo de venda conforme a planilha fornecida.

painel_cabecalho = tk.Frame(
    container,
    bg=PANEL,
    highlightthickness=1,
    highlightbackground=BORDER
)

painel_cabecalho.pack(
    fill="x",
    pady=(0, 10)
)

tk.Label(
    painel_cabecalho,
    text="DADOS DO CABEÇALHO DA VENDA",
    bg=PANEL,
    fg=TEXT,
    font=("Segoe UI", 10, "bold")
).grid(
    row=0,
    column=0,
    columnspan=3,
    sticky="w",
    padx=16,
    pady=(12, 2)
)

tk.Label(
    painel_cabecalho,
    text="Esses campos são utilizados quando o tipo de pedido selecionado é Venda.",
    bg=PANEL,
    fg=MUTED,
    font=("Segoe UI", 8)
).grid(
    row=1,
    column=0,
    columnspan=3,
    sticky="w",
    padx=16,
    pady=(0, 8)
)

for coluna in range(3):
    painel_cabecalho.grid_columnconfigure(
        coluna,
        weight=1
    )


def criar_campo_cabecalho(
    parent,
    titulo,
    row,
    column,
    width=18,
    columnspan=1
):

    bloco = tk.Frame(
        parent,
        bg=PANEL
    )

    bloco.grid(
        row=row,
        column=column,
        columnspan=columnspan,
        sticky="ew",
        padx=16,
        pady=5
    )

    tk.Label(
        bloco,
        text=titulo,
        bg=PANEL,
        fg=MUTED,
        font=("Segoe UI", 8)
    ).pack(anchor="w")

    entrada = tk.Entry(
        bloco,
        width=width,
        bg=PANEL_ALT,
        fg=TEXT,
        insertbackground=TEXT,
        relief="flat",
        font=("Segoe UI", 10),
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT
    )

    entrada.pack(
        fill="x",
        ipady=6,
        pady=(3, 0)
    )

    return entrada


campo_cliente = criar_campo_cabecalho(
    painel_cabecalho,
    "Cliente (até 6 dígitos)",
    2,
    0
)

campo_peso_bruto = criar_campo_cabecalho(
    painel_cabecalho,
    "Peso bruto (até 2 casas decimais)",
    2,
    1
)

campo_volume1 = criar_campo_cabecalho(
    painel_cabecalho,
    "Volume 1 (até 5 dígitos)",
    2,
    2
)

campo_transportadora = criar_campo_cabecalho(
    painel_cabecalho,
    "Transportadora (até 6 dígitos)",
    3,
    0
)

campo_cond_pag = criar_campo_cabecalho(
    painel_cabecalho,
    "Cond. Pag (até 3 dígitos)",
    3,
    1
)

campo_mensagem = criar_campo_cabecalho(
    painel_cabecalho,
    "Mensagem p/ nota",
    4,
    0,
    width=60,
    columnspan=3
)

# Seleção da planilha

painel_arquivo = tk.Frame(
    container,
    bg=PANEL,
    highlightthickness=1,
    highlightbackground=BORDER
)

painel_arquivo.pack(
    fill="x",
    pady=(0, 8)
)

tk.Label(
    painel_arquivo,
    text="PLANILHA",
    bg=PANEL,
    fg=TEXT,
    font=("Segoe UI", 10, "bold")
).pack(
    anchor="w",
    padx=16,
    pady=(10, 5)
)

linha_arquivo = tk.Frame(
    painel_arquivo,
    bg=PANEL
)

linha_arquivo.pack(
    fill="x",
    padx=16,
    pady=(0, 10)
)

ttk.Button(
    linha_arquivo,
    text="Selecionar arquivo Excel",
    style="Secondary.TButton",
    command=selecionar_arquivo
).pack(side="left")

lbl_arquivo = ttk.Label(
    linha_arquivo,
    text="Nenhum arquivo selecionado",
    style="Muted.TLabel"
)

lbl_arquivo.pack(
    side="left",
    padx=12
)

# Ações

linha_acoes = ttk.Frame(container)

linha_acoes.pack(
    fill="x",
    pady=(4, 8)
)

botao_iniciar = ttk.Button(
    linha_acoes,
    text="▶  INICIAR AUTOMAÇÃO",
    style="Accent.TButton",
    command=iniciar_automacao
)

botao_iniciar.pack(side="left")

botao_parar = ttk.Button(
    linha_acoes,
    text="■  PARAR",
    style="Danger.TButton",
    command=parar_automacao,
    state="disabled"
)

botao_parar.pack(
    side="left",
    padx=10
)

ttk.Button(
    linha_acoes,
    text="Limpar log",
    style="Secondary.TButton",
    command=limpar_log
).pack(side="right")

# Log

painel_log = tk.Frame(
    container,
    bg=PANEL,
    highlightthickness=1,
    highlightbackground=BORDER
)

painel_log.pack(
    fill="both",
    expand=True,
    pady=(3, 0)
)

tk.Label(
    painel_log,
    text="LOG DE EXECUÇÃO",
    bg=PANEL,
    fg=TEXT,
    font=("Segoe UI", 10, "bold")
).pack(
    anchor="w",
    padx=16,
    pady=(10, 5)
)

log_text = tk.Text(
    painel_log,
    bg="#0b1220",
    fg="#d1d5db",
    insertbackground=TEXT,
    selectbackground=ACCENT,
    relief="flat",
    borderwidth=0,
    font=("Consolas", 9),
    wrap="word",
    height=6,
    padx=12,
    pady=10,
    state="disabled"
)

log_text.pack(
    fill="both",
    expand=True,
    padx=12,
    pady=(0, 10)
)

ttk.Label(
    container,
    text=(
        "Antes de iniciar, deixe o sistema de pedidos aberto. "
        "Na venda, após preencher o cabeçalho, clique na primeira "
        "linha de produtos."
    ),
    style="Subtitle.TLabel",
    wraplength=920
).pack(
    anchor="w",
    pady=(8, 0)
)


def atualizar_opcoes_por_tipo(*_):

    if tipo_pedido.get() == "Venda":

        check_zerar_lista.configure(
            state="disabled"
        )

        check_preco_unit.configure(
            state="disabled"
        )

        label_preco.configure(
            text=(
                "Venda: preço de lista fixo em 0; "
                "preço unitário lido da coluna K (Custo Trade)."
            )
        )

    else:

        check_zerar_lista.configure(
            state="normal"
        )

        check_preco_unit.configure(
            state="normal"
        )

        label_preco.configure(
            text=(
                "Nas compras, as opções acima controlam "
                "o preenchimento de preço."
            )
        )


combo_tipo.bind(
    "<<ComboboxSelected>>",
    atualizar_opcoes_por_tipo
)

atualizar_opcoes_por_tipo()

janela.mainloop()