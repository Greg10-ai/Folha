import requests
from flask import Flask, render_template, request, session
import pandas as pd
from io import BytesIO
import urllib3
import json
import os
from functools import wraps
from dotenv import load_dotenv
from flask import redirect

load_dotenv()


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.secret_key = 'chave_secreta_para_sessao_sap'

USUARIOS = {
    "admin": "1234",
    "monique": "1998",
    "gregory": "2604",
    "jeferson": "0825",
    "luciana": "Patroa"
}

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logado"):
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated_function
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form.get('usuario')
        senha = request.form.get('senha')

        if usuario in USUARIOS and USUARIOS[usuario] == senha:
            session['logado'] = True
            session['usuario'] = usuario
            return redirect('/')
    else:
        return render_template("login.html", erro="Usuário ou senha inválidos")

    return render_template("login.html")


# =============================
# LOGOUT
# =============================
@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')
# =============================
# CONFIG
# =============================
BPL_ID_PADRAO = 1

# =============================
# ARQUIVO DE PN
# =============================
ARQUIVO_PN = "mapeamento_pn.json"

# =============================
# MAPEAMENTO PN (INICIAL)
# =============================
MAPEAMENTO_PN = {
    "ANA CRISTINA WRUCH": "F001723",
    "ELIZEU FERNANDES DE CARVALHO": "F001924",
    "FLÁVIO MACHADO DOS REIS": "F001603",
    "GISELE PRIEBE PINHEIRO": "F001876",
    "JOSÉ ANTÔNIO OLIVEIRA DA SILVA": "F001820",
    "MATHEUS ALMEIDA DE OLIVEIRA": "F002131",
    "MATHEUS OLLE DE VARGAS": "F001889",
    "JOCIELE NAIARA VAZ PADILHA": "F001997",
    "RAFAEL FERREIRA DA SILVA": "F002020",
    "VANESSA VIEGAS SCHLEMMER": "F001922",
    "WILMA SANTIAGO DA SILVA": "F000746",
    "WILLIAMS SANTOS DE ANDRADE": "F001923"
}

# =============================
# SALVAR / CARREGAR PN
# =============================
def salvar_pn():
    with open(ARQUIVO_PN, "w", encoding="utf-8") as f:
        json.dump(MAPEAMENTO_PN, f, indent=4, ensure_ascii=False)

def carregar_pn():
    global MAPEAMENTO_PN
    if os.path.exists(ARQUIVO_PN):
        try:
            with open(ARQUIVO_PN, "r", encoding="utf-8") as f:
                MAPEAMENTO_PN = json.load(f)
        except:
            pass

carregar_pn()

# =============================
# CONTAS
# =============================
codigos = {
    "4.01.01.01.06": "Valor fixo",
    "4.01.01.06.04": "Comissão variável",
    "4.01.01.01.05": "Comissão fixa",
    "4.01.01.01.14": "Bônus por Tamanho de Carteira",
    "4.01.01.01.17": "Assistência médica",
    "4.01.01.01.18": "Campanhas",
    "4.01.01.01.08": "Bonificação anual",
    "3.01.02.01.12": "Descontos"
}

# =============================
# LEITURA EXCEL
# =============================
def ler_excel_seguro(file_storage):
    file_content = file_storage.read()

    if not file_content:
        raise Exception("Arquivo vazio.")

    file_bytes = BytesIO(file_content)

    for engine in ['openpyxl', 'xlrd', None]:
        try:
            file_bytes.seek(0)
            return pd.read_excel(file_bytes, header=None, engine=engine)
        except:
            continue

    try:
        file_bytes.seek(0)
        return pd.read_csv(file_bytes, header=None, sep=None, engine='python')
    except:
        pass

    raise Exception("Arquivo inválido.")

# =============================
# LOGIN SAP
# =============================
def login_sap():
    url = "https://b1.ativy.com:51032/b1s/v1/Login"

    payload = {
        "CompanyDB": os.getenv("CompanyDB"),
        "UserName": os.getenv("SAP_USER"),
        "Password": os.getenv("SAP_PASSWORD")
    }

    r = requests.post(url, json=payload, verify=False)

    print("LOGIN:", r.status_code, r.text)

    return r.cookies if r.status_code == 200 else None

# =============================
# PROCESSAMENTO
# =============================
def processar_dados(file_rel, mes, dt_l, dt_v, vendedor_alvo):
    df = ler_excel_seguro(file_rel)

    linha_meses = df.iloc[0, :].ffill().tolist()
    vendedores_row = df.iloc[1, :].tolist()

    linhas = []
    CONTA_DEDUCAO = "3.01.02.01.12"

    for col in range(2, df.shape[1]):

        mes_col = str(linha_meses[col]).upper().strip()

        if mes.upper() not in mes_col:
            continue

        vendedor_nome = str(vendedores_row[col]).replace("C.", "").strip().upper()

        if pd.isna(vendedores_row[col]) or "TOTAL" in vendedor_nome:
            continue

        if vendedor_alvo != "TODOS" and vendedor_nome != vendedor_alvo:
            continue

        soma = 0

        for row in range(2, df.shape[0]):

            conta = str(df.iloc[row, 0]).strip()
            valor = df.iloc[row, col]

            if pd.notnull(valor) and valor != 0 and conta != "nan":

                valor = float(valor)
                is_deducao = conta == CONTA_DEDUCAO

                soma += -abs(valor) if is_deducao else valor

                codigo_pn = MAPEAMENTO_PN.get(vendedor_nome)

                if not codigo_pn:
                    raise Exception(f"PN não encontrado: {vendedor_nome}")

                nome_conta = codigos.get(conta, "")

                linhas.append({
                    "AccountCode": conta,
                    "Debit": abs(valor) if valor > 0 and not is_deducao else 0,
                    "Credit": abs(valor) if valor < 0 or is_deducao else 0,
                    "LineMemo": f"{vendedor_nome} - {nome_conta}",
                    "BPLID": BPL_ID_PADRAO
                })

        if soma != 0:
            codigo_pn = MAPEAMENTO_PN.get(vendedor_nome)

            linhas.append({
                "ShortName": codigo_pn,
                "Debit": 0,
                "Credit": abs(soma),
                "LineMemo": f"TOTAL LIQUIDO - {vendedor_nome}",
                "BPLID": BPL_ID_PADRAO
            })

    return {
        "ReferenceDate": dt_l,
        "DueDate": dt_v,
        "TaxDate": dt_l,
        "Memo": f"FECHAMENTO PJ - {mes}",
        "JournalEntryLines": linhas
    }

# =============================
# NOVA ROTA - ADD VENDEDOR
# =============================
@app.route('/add_vendedor', methods=['POST'])
def add_vendedor():
    try:
        nome = request.form.get('nome').strip().upper()
        pn = request.form.get('pn').strip()

        if not nome or not pn:
            return render_template("retorno.html", mensagem="❌ Nome e PN são obrigatórios")

        if nome in MAPEAMENTO_PN:
            return render_template("retorno.html", mensagem=f"⚠️ Vendedor já existe: {nome}")

        MAPEAMENTO_PN[nome] = pn
        salvar_pn()

        return render_template("retorno.html", mensagem="✅ Vendedor adicionado com sucesso")

    except Exception as e:
        return render_template("retorno.html", mensagem=f"❌ Erro: {str(e)}")
    
# =============================
# ATUALIZAR PN LOCAL
# =============================
@app.route('/editar_vendedor', methods=['POST'])
def editar_vendedor():
    try:
        nome = request.form.get('nome').strip().upper()
        novo_pn = request.form.get('pn').strip()

        if nome not in MAPEAMENTO_PN:
            return render_template("retorno.html",mensagem="❌ Vendedor não encontrado")

        MAPEAMENTO_PN[nome] = novo_pn

        salvar_pn()

        return render_template("retorno.html", mensagem="✅ PN atualizado com sucesso")

    except Exception as e:
        return render_template("retorno.html",
                               mensagem=f"❌ Erro: {str(e)}")

# =============================
# ROTAS
# =============================
@app.route('/')
@login_required
def index():
    meses = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
             "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]

    return render_template("index.html",
                           meses=meses,
                           vendedores_lista=sorted(MAPEAMENTO_PN.keys()))

@app.route('/processar', methods=['POST'])
@login_required
def processar():
    try:
        acao = request.form.get('acao')

        if acao == 'postar':

            json_editado = request.form.get('json_editado')
            sap_json = json.loads(json_editado) if json_editado else session.get('last_sap_json')

            if not sap_json:
                return "❌ Dados expirados"

            cookies = login_sap()
            if not cookies:
                return "❌ Erro login SAP"

            r = requests.post(
                "https://b1.ativy.com:51032/b1s/v1/JournalEntries",
                json=sap_json,
                cookies=cookies,
                headers={"Content-Type": "application/json"},
                verify=False
            )

            if r.status_code in [200, 201]:
                resp = r.json()
                doc = resp.get('JdtNum') or resp.get('TransId')
                return f"✅ Lançamento criado: {doc}"

            return f"❌ Erro SAP: {r.text}"

        rel = request.files['relatorio']

        payload = processar_dados(
            rel,
            request.form.get('mes'),
            request.form.get('data_lancamento'),
            request.form.get('data_vencimento'),
            request.form.get('vendedor')
        )

        session['last_sap_json'] = payload

        return render_template("index.html",
                               meses=["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
                                      "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"],
                               vendedores_lista=sorted(MAPEAMENTO_PN.keys()),
                               payload=payload,
                               show_post=True)

    except Exception as e:
        return f"❌ Erro: {str(e)}"

# =============================
# START
# =============================
if __name__ == "__main__":
    print("🚀 Servidor iniciando...")

    port = int(os.environ.get("PORT", 5000))

    print(f"🌐 Rodando na porta: {port}")
    print(f"🔗 Acesse: http://localhost:{port}/login")

    app.run(host="0.0.0.0", port=port)
