"""
===================================================================
 Emissor de Certificados de Regularidade Associativa
 Mercado Central de Belo Horizonte
 -----------------------------------------------------------------
 Desenvolvido por: Artur Gabriel Oliveira da Silva
 Contato:          arturg_oliveira@outlook.com
===================================================================

 Funcionalidades
 ---------------
 • Login com PIN (hash SHA256 + salt)
 • Carregamento de base de associados via planilha Excel
 • Busca com autocomplete inteligente (por nome ou matrícula)
 • Preenchimento automático dos dados ao selecionar associado
 • Campos editáveis para correção pontual antes do envio
 • Geração do certificado em PDF (docxtpl + docx2pdf)
 • QR Code de autenticidade embutido em cada certificado
 • Envio por e-mail (SMTP + anexo)
 • Envio por WhatsApp (wa.me + PDF pronto para anexar)
 • Log de auditoria em CSV
 • Validação de inputs (e-mail, telefone)
 • Credenciais SMTP isoladas em .env
 • Bloqueio após 3 tentativas de PIN incorreto
"""

import os
import re
import sys
import csv
import json
import hmac
import uuid
import hashlib
import secrets
import logging
import smtplib
import webbrowser
import subprocess
import unicodedata
from pathlib import Path
from datetime import datetime, timedelta
from urllib.parse import quote
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Dependências externas
try:
    import pandas as pd
    from docxtpl import DocxTemplate, InlineImage
    from docx.shared import Mm
    import qrcode
    from dotenv import load_dotenv
except ImportError as e:
    missing = str(e).split("'")[1] if "'" in str(e) else str(e)
    print(f"ERRO: dependência faltando: {missing}")
    print("Instale com: pip install -r requirements.txt")
    sys.exit(1)


# ============================================================
# CONFIGURAÇÃO
# ============================================================
BASE_DIR       = Path(__file__).resolve().parent
ASSETS_DIR     = BASE_DIR
OUTPUT_DIR     = BASE_DIR / "certificados_gerados"
LOG_DIR        = BASE_DIR / "logs"
CONFIG_FILE    = BASE_DIR / "config.json"
LOGO_PATH      = BASE_DIR / "logo.png"
MODELO_PATH    = BASE_DIR / "modelo_certificado.docx"
PLANILHA_PATH  = BASE_DIR / "planilha.xlsx"
ENV_PATH       = BASE_DIR / ".env"

OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

load_dotenv(ENV_PATH)

SMTP_HOST   = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER   = os.getenv("SMTP_USER", "")
SMTP_PASS   = os.getenv("SMTP_PASS", "")
SMTP_FROM   = os.getenv("SMTP_FROM", "Mercado Central BH")

# Paleta (identidade Mercado Central)
NAVY      = "#1A2332"
NAVY_DARK = "#0F1A28"
GOLD      = "#C9A961"
GOLD_DEEP = "#9C7F3E"
IVORY     = "#F8F5F0"
SLATE     = "#4A5568"
SLATE_LT  = "#718096"
DIVIDER   = "#E2DDD3"
SUCCESS   = "#2F855A"
DANGER    = "#C53030"
WHITE     = "#FFFFFF"

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"app_{datetime.now():%Y%m%d}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("app")

AUDIT_CSV = LOG_DIR / "auditoria_envios.csv"
if not AUDIT_CSV.exists():
    with open(AUDIT_CSV, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(
            ["timestamp", "operador", "matricula", "nome", "destino", "canal", "status", "detalhe"]
        )

HISTORICO_CSV = LOG_DIR / "historico_emissoes.csv"


def registrar_auditoria(operador, matricula, nome, destino, canal, status, detalhe=""):
    """Registra cada ação no CSV de auditoria (compliance LGPD)."""
    try:
        with open(AUDIT_CSV, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                datetime.now().isoformat(timespec="seconds"),
                operador, matricula, nome, destino, canal, status, detalhe
            ])
    except Exception as e:
        log.warning(f"Falha ao registrar auditoria: {e}")


# ============================================================
# SEGURANÇA  (hash de PIN + controle de tentativas)
# ============================================================
class Seguranca:
    """Gerencia PIN de acesso, bloqueios e hash seguro."""

    MAX_TENTATIVAS = 3
    BLOQUEIO_MIN   = 5

    def __init__(self):
        self.config = self._carregar_config()
        self.tentativas = 0
        self.bloqueado_ate = None

    # ---------- config ----------
    def _carregar_config(self):
        if CONFIG_FILE.exists():
            try:
                return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            except Exception:
                log.warning("config.json corrompido; recriando.")
        return {}

    def _salvar_config(self):
        CONFIG_FILE.write_text(json.dumps(self.config, indent=2), encoding="utf-8")

    # ---------- PIN ----------
    @staticmethod
    def _hash_pin(pin: str, salt: str) -> str:
        """PBKDF2-HMAC-SHA256, 100k iterações."""
        return hashlib.pbkdf2_hmac(
            "sha256", pin.encode("utf-8"), salt.encode("utf-8"), 100_000
        ).hex()

    def pin_configurado(self) -> bool:
        return "pin_hash" in self.config and "pin_salt" in self.config

    def definir_pin(self, pin: str, operador: str, email_recuperacao: str = ""):
        salt = secrets.token_hex(16)
        self.config["pin_hash"] = self._hash_pin(pin, salt)
        self.config["pin_salt"] = salt
        self.config["operador"] = operador
        self.config["email_recuperacao"] = email_recuperacao.strip()
        self.config["criado_em"] = datetime.now().isoformat()
        self._salvar_config()
        log.info(f"PIN configurado para operador '{operador}'.")

    def trocar_pin(self, pin_atual: str, pin_novo: str) -> tuple[bool, str]:
        """Valida PIN atual e atualiza para novo. Retorna (ok, mensagem)."""
        if not self.verificar_pin(pin_atual):
            return False, "PIN atual incorreto."
        if not (4 <= len(pin_novo) <= 8) or not pin_novo.isdigit():
            return False, "Novo PIN deve ter entre 4 e 8 dígitos numéricos."
        salt = secrets.token_hex(16)
        self.config["pin_hash"] = self._hash_pin(pin_novo, salt)
        self.config["pin_salt"] = salt
        self.config["pin_alterado_em"] = datetime.now().isoformat()
        self._salvar_config()
        self.tentativas = 0
        log.info("PIN alterado com sucesso.")
        return True, "PIN alterado com sucesso."

    def email_recuperacao(self) -> str:
        return self.config.get("email_recuperacao", "")

    def definir_email_recuperacao(self, email: str):
        self.config["email_recuperacao"] = email.strip()
        self._salvar_config()

    # ---------- Token temporário de recuperação ----------
    def gerar_token_recuperacao(self) -> str:
        """Gera código de 6 dígitos válido por 15 minutos."""
        import random
        codigo = f"{random.randint(100000, 999999)}"
        self.config["recuperacao_token"] = self._hash_pin(codigo, "recuperacao_salt")
        self.config["recuperacao_expira"] = (
            datetime.now() + timedelta(minutes=15)
        ).isoformat()
        self._salvar_config()
        return codigo

    def validar_token_recuperacao(self, codigo: str) -> bool:
        """Verifica se o código de recuperação é válido e não expirou."""
        token_hash = self.config.get("recuperacao_token", "")
        expira_str = self.config.get("recuperacao_expira", "")
        if not token_hash or not expira_str:
            return False
        try:
            expira = datetime.fromisoformat(expira_str)
            if datetime.now() > expira:
                return False
        except Exception:
            return False
        return hmac.compare_digest(
            token_hash, self._hash_pin(codigo, "recuperacao_salt")
        )

    def limpar_token_recuperacao(self):
        """Remove token de recuperação após uso."""
        for k in ("recuperacao_token", "recuperacao_expira"):
            self.config.pop(k, None)
        self._salvar_config()

    def redefinir_pin_apos_recuperacao(self, pin_novo: str):
        """Define novo PIN após validação por token. Limpa bloqueio."""
        salt = secrets.token_hex(16)
        self.config["pin_hash"] = self._hash_pin(pin_novo, salt)
        self.config["pin_salt"] = salt
        self.config["pin_alterado_em"] = datetime.now().isoformat()
        self._salvar_config()
        self.tentativas = 0
        self.bloqueado_ate = None
        self.limpar_token_recuperacao()
        log.info("PIN redefinido via recuperação por e-mail.")

    def verificar_pin(self, pin: str) -> bool:
        if self.esta_bloqueado():
            return False
        ok = hmac.compare_digest(
            self.config.get("pin_hash", ""),
            self._hash_pin(pin, self.config.get("pin_salt", ""))
        )
        if ok:
            self.tentativas = 0
            return True
        self.tentativas += 1
        if self.tentativas >= self.MAX_TENTATIVAS:
            self.bloqueado_ate = datetime.now() + timedelta(minutes=self.BLOQUEIO_MIN)
            log.warning("Sistema bloqueado por excesso de tentativas.")
        return False

    def esta_bloqueado(self) -> bool:
        if self.bloqueado_ate and datetime.now() < self.bloqueado_ate:
            return True
        if self.bloqueado_ate and datetime.now() >= self.bloqueado_ate:
            self.bloqueado_ate = None
            self.tentativas = 0
        return False

    def tempo_bloqueio_restante(self) -> int:
        if not self.bloqueado_ate:
            return 0
        restante = (self.bloqueado_ate - datetime.now()).total_seconds()
        return max(0, int(restante))

    def operador(self) -> str:
        return self.config.get("operador", "desconhecido")

    # ---------- Signatário e dados do certificado ----------
    def signatario_nome(self) -> str:
        return self.config.get("signatario_nome", "Jose Agostinho Oliveira Quadros")

    def signatario_cargo(self) -> str:
        return self.config.get("signatario_cargo", "Diretor Secretário")

    def signatario_assinatura_path(self) -> str | None:
        path = self.config.get("signatario_assinatura_path")
        if path and Path(path).exists():
            return path
        return None

    def data_assembleia_padrao(self) -> str:
        return self.config.get("data_assembleia_padrao", "27 de abril de 2026")

    def atualizar_signatario(self, nome: str, cargo: str,
                              assinatura_path: str | None,
                              data_assembleia: str):
        self.config["signatario_nome"] = nome.strip()
        self.config["signatario_cargo"] = cargo.strip()
        if assinatura_path:
            self.config["signatario_assinatura_path"] = str(assinatura_path)
        elif "signatario_assinatura_path" in self.config:
            del self.config["signatario_assinatura_path"]
        self.config["data_assembleia_padrao"] = data_assembleia.strip()
        self._salvar_config()
        log.info(f"Configurações do signatário atualizadas: {nome}")

    def atualizar_operador(self, nome: str):
        self.config["operador"] = nome.strip()
        self._salvar_config()


# ============================================================
# VALIDAÇÃO E SANITIZAÇÃO
# ============================================================
REGEX_EMAIL = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

def validar_email(email: str) -> bool:
    return bool(email and REGEX_EMAIL.match(email.strip()))

def normalizar_telefone(telefone) -> str | None:
    """Normaliza para 55DDDNNNNNNNNN (formato wa.me)."""
    if telefone is None:
        return None
    dig = "".join(filter(str.isdigit, str(telefone)))
    if len(dig) == 11:          # DDD + 9 dígitos
        return "55" + dig
    if len(dig) == 10:          # DDD + 8 dígitos (fixo antigo)
        return "55" + dig
    if len(dig) == 13 and dig.startswith("55"):
        return dig
    if len(dig) == 12 and dig.startswith("55"):
        return dig
    return None

def sanitizar_texto(texto: str) -> str:
    """Remove chars que podem quebrar o template Jinja2 do docxtpl."""
    if not texto:
        return ""
    # Remove { } para evitar injeção de tags
    return str(texto).replace("{", "").replace("}", "").strip()

def remover_acentos(texto: str) -> str:
    """Normaliza texto removendo acentos (para busca case-insensitive)."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


# ============================================================
# REPOSITÓRIO DE ASSOCIADOS
# ============================================================
class RepositorioAssociados:
    """Carrega planilha Excel e oferece busca com autocomplete.

    Formato flexível: aceita planilhas com apenas `nome` e `matricula`,
    onde email e telefone podem estar ausentes (serão preenchidos
    pelo operador no momento da emissão).

    Também detecta automaticamente variações nos nomes de colunas:
      - matricula: "matricula", "matrícula", "matric", "id"
      - nome: "nome", "associado", "nome completo"
      - email: "email", "e-mail", "e_mail"
      - telefone: "telefone", "celular", "whatsapp", "tel"
    """

    # Mapeamento de sinônimos → nome canônico
    ALIASES = {
        "matricula":  {"matricula", "matrícula", "matric", "matr", "id", "codigo", "código"},
        "nome":       {"nome", "associado", "nome completo", "nomecompleto", "nome_completo"},
        "email":      {"email", "e-mail", "e_mail", "email_pessoal", "emailpessoal"},
        "telefone":   {"telefone", "celular", "whatsapp", "tel", "fone", "whats", "telefone/whatsapp"},
    }

    COLUNAS_MINIMAS = {"matricula", "nome"}

    def __init__(self):
        self.registros = []       # [{matricula, nome, email, telefone, _busca}]
        self.carregado = False
        self.fonte = None

    @classmethod
    def _normalizar_colunas(cls, colunas: list[str]) -> dict:
        """Mapeia colunas da planilha para os nomes canônicos. Retorna {canonico: col_original}."""
        mapa = {}
        for col in colunas:
            col_limpa = col.strip().lower().replace(" ", "").replace("-", "").replace("_", "")
            for canonico, aliases in cls.ALIASES.items():
                aliases_limpos = {a.replace(" ", "").replace("-", "").replace("_", "") for a in aliases}
                if col_limpa in aliases_limpos:
                    mapa[canonico] = col
                    break
        return mapa

    def carregar(self, path: Path) -> tuple[bool, str]:
        try:
            df = pd.read_excel(path, dtype=str).fillna("")
            df.columns = df.columns.str.strip()

            # Detectar colunas por aliases (flexível)
            mapa = self._normalizar_colunas(list(df.columns))

            # Colunas mínimas: matricula + nome
            faltando = self.COLUNAS_MINIMAS - set(mapa.keys())
            if faltando:
                return False, (
                    f"Planilha não tem as colunas necessárias.\n"
                    f"Colunas mínimas: matricula e nome.\n"
                    f"Colunas encontradas: {', '.join(df.columns)}"
                )

            self.registros = []
            for _, row in df.iterrows():
                nome = str(row[mapa["nome"]]).strip()
                matr = str(row[mapa["matricula"]]).strip()
                if not nome and not matr:
                    continue

                # Email e telefone são OPCIONAIS
                email = ""
                tel = ""
                if "email" in mapa:
                    email = str(row[mapa["email"]]).strip()
                if "telefone" in mapa:
                    tel = str(row[mapa["telefone"]]).strip()

                self.registros.append({
                    "matricula": matr,
                    "nome":      nome,
                    "email":     email,
                    "telefone":  tel,
                    "_busca":    remover_acentos(f"{nome} {matr}")
                })
            self.carregado = True
            self.fonte = str(path)

            # Mensagem informativa sobre colunas detectadas
            info_cols = [f"{k}='{v}'" for k, v in mapa.items()]
            log.info(f"Planilha carregada: {len(self.registros)} registros. "
                     f"Colunas detectadas: {', '.join(info_cols)}")

            tem_email_tel = "email" in mapa and "telefone" in mapa
            aviso = ""
            if not tem_email_tel:
                faltam = []
                if "email" not in mapa: faltam.append("e-mail")
                if "telefone" not in mapa: faltam.append("telefone")
                aviso = (f" ({', '.join(faltam)} não encontrado(s); "
                         f"será(ão) preenchido(s) manualmente)")

            return True, f"{len(self.registros)} associados carregados{aviso}"
        except Exception as e:
            log.exception("Erro ao carregar planilha")
            return False, str(e)

    def buscar(self, termo: str, limite: int = 8) -> list[dict]:
        """Busca fuzzy por nome ou matrícula. Retorna até `limite` matches."""
        if not self.carregado or not termo or len(termo.strip()) < 2:
            return []
        alvo = remover_acentos(termo)
        tokens = alvo.split()
        resultados = []
        for reg in self.registros:
            # Todos os tokens precisam estar presentes em algum lugar
            if all(t in reg["_busca"] for t in tokens):
                resultados.append(reg)
                if len(resultados) >= limite:
                    break
        return resultados

    def total(self) -> int:
        return len(self.registros)


# ============================================================
# GERAÇÃO DE CERTIFICADO
# ============================================================
class GeradorCertificado:

    @staticmethod
    def _gerar_codigo_verificacao(matricula: str) -> str:
        """Código curto e único para autenticidade (ex: MC-2026-A4F7B2)."""
        seed = f"{matricula}-{datetime.now().isoformat()}-{uuid.uuid4().hex}"
        h = hashlib.sha256(seed.encode()).hexdigest()[:8].upper()
        return f"MC-{datetime.now().year}-{h}"

    @staticmethod
    def _gerar_qr(codigo: str, nome: str, matricula: str, saida: Path):
        """Gera imagem PNG do QR Code contendo dados de verificação."""
        payload = (
            f"CERTIFICADO MERCADO CENTRAL BH\n"
            f"Codigo: {codigo}\n"
            f"Nome: {nome}\n"
            f"Matricula: {matricula}\n"
            f"Emitido: {datetime.now():%d/%m/%Y %H:%M}"
        )
        img = qrcode.make(payload)
        img.save(saida)
        return saida

    def gerar(self, nome: str, matricula: str, data_assembleia: str,
              signatario_nome: str = "Jose Agostinho Oliveira Quadros",
              signatario_cargo: str = "Diretor Secretário",
              assinatura_path: str | None = None
              ) -> tuple[Path | None, str, str]:
        """
        Gera o PDF do certificado.
        Retorna: (caminho_pdf, codigo_verificacao, erro_ou_vazio)
        """
        try:
            nome = sanitizar_texto(nome)
            matricula = sanitizar_texto(matricula)

            codigo = self._gerar_codigo_verificacao(matricula)
            nome_safe = re.sub(r"[^\w\s-]", "", nome).strip().replace(" ", "_")[:40]
            base = f"Certificado_{matricula}_{nome_safe}_{datetime.now():%H%M%S}"

            qr_path  = OUTPUT_DIR / f".qr_{base}.png"
            docx_out = OUTPUT_DIR / f"{base}.docx"
            pdf_out  = OUTPUT_DIR / f"{base}.pdf"

            self._gerar_qr(codigo, nome, matricula, qr_path)

            doc = DocxTemplate(MODELO_PATH)

            # Assinatura: imagem se configurada, linha em branco se não
            if assinatura_path and Path(assinatura_path).exists():
                assinatura_valor = InlineImage(doc, str(assinatura_path), width=Mm(55))
            else:
                assinatura_valor = "_" * 50  # linha tradicional

            ctx = {
                "nome":                nome.upper(),
                "matricula":           matricula,
                "data_assembleia":     data_assembleia,
                "data_emissao":        datetime.now().strftime("%d de %B de %Y").lower()
                                           .replace("january", "janeiro")
                                           .replace("february", "fevereiro")
                                           .replace("march", "março")
                                           .replace("april", "abril")
                                           .replace("may", "maio")
                                           .replace("june", "junho")
                                           .replace("july", "julho")
                                           .replace("august", "agosto")
                                           .replace("september", "setembro")
                                           .replace("october", "outubro")
                                           .replace("november", "novembro")
                                           .replace("december", "dezembro"),
                "codigo_verificacao":  codigo,
                "qr_code":             InlineImage(doc, str(qr_path), width=Mm(25)),
                "signatario_nome":     sanitizar_texto(signatario_nome),
                "signatario_cargo":    sanitizar_texto(signatario_cargo),
                "assinatura_img":      assinatura_valor,
            }
            doc.render(ctx)
            doc.save(docx_out)

            # ---- Conversão para PDF ----
            if self._converter_para_pdf(docx_out, pdf_out):
                # Limpar arquivos intermediários
                docx_out.unlink(missing_ok=True)
                qr_path.unlink(missing_ok=True)
                return pdf_out, codigo, ""
            else:
                # Se falhou a conversão, retorna o .docx mesmo
                qr_path.unlink(missing_ok=True)
                return docx_out, codigo, "PDF não pôde ser gerado; entregando DOCX."

        except Exception as e:
            log.exception("Erro ao gerar certificado")
            return None, "", str(e)

    def _converter_para_pdf(self, docx: Path, pdf: Path) -> bool:
        """Tenta via docx2pdf (Windows/Mac) → fallback LibreOffice."""
        # Tentativa 1: docx2pdf (usa MS Word no Windows)
        try:
            if sys.platform == "win32":
                import pythoncom
                pythoncom.CoInitialize()
            from docx2pdf import convert
            convert(str(docx), str(pdf))
            if sys.platform == "win32":
                pythoncom.CoUninitialize()
            if pdf.exists():
                return True
        except Exception as e:
            log.warning(f"docx2pdf falhou: {e}")

        # Tentativa 2: LibreOffice headless
        try:
            result = subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf",
                 "--outdir", str(pdf.parent), str(docx)],
                capture_output=True, timeout=30, text=True
            )
            if pdf.exists():
                return True
            log.warning(f"LibreOffice stderr: {result.stderr}")
        except Exception as e:
            log.warning(f"LibreOffice falhou: {e}")
        return False


# ============================================================
# ENVIO  (E-mail + WhatsApp)
# ============================================================
class Envio:

    @staticmethod
    def email(destinatario: str, arquivo: Path, nome_pessoa: str,
              matricula: str, operador: str) -> tuple[bool, str]:
        if not validar_email(destinatario):
            registrar_auditoria(operador, matricula, nome_pessoa, destinatario,
                                "email", "FALHA", "email invalido")
            return False, "E-mail inválido."
        if not SMTP_USER or not SMTP_PASS:
            return False, "Credenciais SMTP não configuradas no .env"

        try:
            msg = MIMEMultipart()
            msg["From"] = f"{SMTP_FROM} <{SMTP_USER}>"
            msg["To"] = destinatario
            msg["Subject"] = "Certificado de Regularidade Associativa - Mercado Central BH"

            corpo = (
                f"Prezado(a) {nome_pessoa},\n\n"
                f"Segue em anexo o seu Certificado de Regularidade Associativa, "
                f"referente à Assembleia Geral Ordinária do Mercado Central de "
                f"Belo Horizonte.\n\n"
                f"Matrícula: {matricula}\n"
                f"Data de emissão: {datetime.now():%d/%m/%Y}\n\n"
                f"Em caso de dúvidas, responda este e-mail.\n\n"
                f"Atenciosamente,\n"
                f"Secretaria do Mercado Central de Belo Horizonte\n"
            )
            msg.attach(MIMEText(corpo, "plain", "utf-8"))

            with open(arquivo, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="{arquivo.name}"')
            msg.attach(part)

            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)

            registrar_auditoria(operador, matricula, nome_pessoa, destinatario,
                                "email", "SUCESSO", arquivo.name)
            log.info(f"E-mail enviado: {destinatario}")
            return True, f"E-mail enviado com sucesso para {destinatario}"

        except smtplib.SMTPAuthenticationError:
            msg = "Falha de autenticação SMTP. Verifique usuário/senha no .env"
            registrar_auditoria(operador, matricula, nome_pessoa, destinatario,
                                "email", "FALHA", "auth SMTP")
            return False, msg
        except Exception as e:
            registrar_auditoria(operador, matricula, nome_pessoa, destinatario,
                                "email", "FALHA", str(e)[:100])
            log.exception("Erro no envio de e-mail")
            return False, f"Erro: {e}"

    @staticmethod
    def whatsapp(telefone, arquivo: Path, nome: str, matricula: str,
                 operador: str) -> tuple[bool, str]:
        numero = normalizar_telefone(telefone)
        if not numero:
            registrar_auditoria(operador, matricula, nome, str(telefone),
                                "whatsapp", "FALHA", "telefone invalido")
            return False, f"Telefone inválido: {telefone}"

        mensagem = (
            f"Olá {nome}!\n\n"
            f"Segue em anexo o seu Certificado de Regularidade Associativa "
            f"do Mercado Central de Belo Horizonte.\n\n"
            f"Matrícula: {matricula}\n"
            f"Data de emissão: {datetime.now():%d/%m/%Y}"
        )
        link = f"https://wa.me/{numero}?text={quote(mensagem)}"
        webbrowser.open(link)

        # Abre a pasta do certificado para o operador anexar manualmente
        try:
            if sys.platform == "win32":
                subprocess.Popen(["explorer", "/select,", str(arquivo)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", str(arquivo)])
            else:
                subprocess.Popen(["xdg-open", str(arquivo.parent)])
        except Exception as e:
            log.warning(f"Não consegui abrir a pasta: {e}")

        registrar_auditoria(operador, matricula, nome, numero, "whatsapp",
                            "LINK_ABERTO", arquivo.name)
        return True, ("WhatsApp Web aberto. A pasta do certificado também foi "
                      "aberta — arraste o PDF para o chat.")

    @staticmethod
    def email_simples(destinatario: str, assunto: str, corpo: str) -> tuple[bool, str]:
        """Envia um e-mail sem anexo (usado para recuperação de PIN)."""
        if not validar_email(destinatario):
            return False, "E-mail inválido."
        if not SMTP_USER or not SMTP_PASS:
            return False, "Credenciais SMTP não configuradas no arquivo .env"
        try:
            msg = MIMEMultipart()
            msg["From"] = f"{SMTP_FROM} <{SMTP_USER}>"
            msg["To"] = destinatario
            msg["Subject"] = assunto
            msg.attach(MIMEText(corpo, "plain", "utf-8"))

            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)
            log.info(f"E-mail simples enviado: {destinatario}")
            return True, "E-mail enviado."
        except smtplib.SMTPAuthenticationError:
            return False, "Falha de autenticação SMTP. Verifique o .env"
        except Exception as e:
            log.exception("Erro no envio de e-mail simples")
            return False, f"Erro: {e}"


# ============================================================
# HISTÓRICO DE EMISSÕES
# ============================================================
class HistoricoEmissoes:
    """
    Gerencia o histórico persistente de certificados emitidos.
    Armazenado em logs/historico.csv para sobreviver entre sessões.
    """

    CAMPOS = [
        "timestamp", "codigo", "matricula", "nome", "email", "telefone",
        "data_assembleia", "arquivo_pdf", "operador", "enviado_email", "enviado_whatsapp"
    ]

    def __init__(self, path: Path):
        self.path = path
        if not path.exists():
            with open(path, "w", newline="", encoding="utf-8") as f:
                csv.DictWriter(f, fieldnames=self.CAMPOS).writeheader()

    def registrar(self, dados: dict):
        """Adiciona emissão ao histórico."""
        registro = {k: dados.get(k, "") for k in self.CAMPOS}
        registro["timestamp"] = datetime.now().isoformat(timespec="seconds")
        try:
            with open(self.path, "a", newline="", encoding="utf-8") as f:
                csv.DictWriter(f, fieldnames=self.CAMPOS).writerow(registro)
        except Exception as e:
            log.warning(f"Falha ao registrar histórico: {e}")

    def marcar_enviado(self, codigo: str, canal: str):
        """Atualiza flag enviado_email ou enviado_whatsapp no registro."""
        if canal not in ("email", "whatsapp"):
            return
        campo = f"enviado_{canal}"
        try:
            # Ler tudo, atualizar, reescrever
            with open(self.path, "r", encoding="utf-8", newline="") as f:
                rows = list(csv.DictReader(f))
            for r in rows:
                if r.get("codigo") == codigo:
                    r[campo] = "S"
            with open(self.path, "w", encoding="utf-8", newline="") as f:
                w = csv.DictWriter(f, fieldnames=self.CAMPOS)
                w.writeheader()
                w.writerows(rows)
        except Exception as e:
            log.warning(f"Falha ao atualizar histórico: {e}")

    def listar(self, filtro_texto: str = "", data_de: str = "",
               data_ate: str = "") -> list[dict]:
        """Retorna emissões filtradas. Mais recentes primeiro."""
        try:
            with open(self.path, "r", encoding="utf-8", newline="") as f:
                rows = list(csv.DictReader(f))
        except Exception:
            return []

        filtro = remover_acentos(filtro_texto.strip())

        def _passa(r):
            if filtro:
                alvo = remover_acentos(f"{r.get('nome','')} {r.get('matricula','')} {r.get('codigo','')}")
                if filtro not in alvo:
                    return False
            ts = r.get("timestamp", "")[:10]
            if data_de and ts < data_de:
                return False
            if data_ate and ts > data_ate:
                return False
            return True

        return sorted([r for r in rows if _passa(r)],
                      key=lambda r: r.get("timestamp", ""), reverse=True)

    def buscar_por_codigo(self, codigo: str) -> dict | None:
        try:
            with open(self.path, "r", encoding="utf-8", newline="") as f:
                for r in csv.DictReader(f):
                    if r.get("codigo") == codigo:
                        return r
        except Exception:
            pass
        return None


# ============================================================
# WIDGET DE AUTOCOMPLETE
# ============================================================
class AutocompleteEntry(tk.Frame):
    """Entry com dropdown de sugestões."""

    def __init__(self, parent, repo: RepositorioAssociados, on_select, **kwargs):
        super().__init__(parent, bg=kwargs.pop("bg", WHITE))
        self.repo = repo
        self.on_select = on_select
        self._ignore_next = False

        self.var = tk.StringVar()
        self.entry = tk.Entry(
            self, textvariable=self.var,
            font=("Calibri", 13), relief="solid", bd=1,
            highlightthickness=1, highlightcolor=GOLD,
            highlightbackground=DIVIDER
        )
        self.entry.pack(fill="x", ipady=6)

        self.var.trace_add("write", self._on_type)
        self.entry.bind("<Down>",   self._focus_list)
        self.entry.bind("<Return>", self._enter_key)
        self.entry.bind("<Escape>", lambda e: self._hide_list())
        self.entry.bind("<FocusOut>", self._delayed_hide)

        # Listbox em Toplevel flutuante (segue o Entry)
        self.top = None
        self.listbox = None

    # ---------- API ----------
    def get(self) -> str:
        return self.var.get()

    def set(self, texto: str):
        self._ignore_next = True
        self.var.set(texto)

    def focus(self):
        self.entry.focus_set()

    # ---------- internos ----------
    def _on_type(self, *_):
        if self._ignore_next:
            self._ignore_next = False
            return
        termo = self.var.get()
        if len(termo.strip()) < 2:
            self._hide_list()
            return
        matches = self.repo.buscar(termo)
        if matches:
            self._show_list(matches)
        else:
            self._hide_list()

    def _show_list(self, matches):
        if self.top is None:
            self.top = tk.Toplevel(self)
            self.top.wm_overrideredirect(True)
            self.top.configure(bg=NAVY)
            self.listbox = tk.Listbox(
                self.top, font=("Calibri", 11),
                bg=WHITE, fg=SLATE,
                selectbackground=NAVY, selectforeground=WHITE,
                activestyle="none", relief="flat", bd=0,
                highlightthickness=0, height=min(8, len(matches))
            )
            self.listbox.pack(fill="both", expand=True, padx=1, pady=1)
            self.listbox.bind("<Double-Button-1>", self._select_item)
            self.listbox.bind("<Return>",           self._select_item)
            self.listbox.bind("<Escape>",           lambda e: self._hide_list())
            self.listbox.bind("<FocusOut>",         self._delayed_hide)

        self.listbox.delete(0, tk.END)
        self._matches = matches
        for m in matches:
            display = f"  {m['nome']}   │   Matr. {m['matricula']}"
            self.listbox.insert(tk.END, display)

        # Posiciona logo abaixo do Entry
        self.update_idletasks()
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        w = self.entry.winfo_width()
        h = min(8, len(matches)) * 22 + 6
        self.top.geometry(f"{w}x{h}+{x}+{y}")
        self.top.deiconify()
        self.top.lift()

    def _hide_list(self):
        if self.top is not None:
            try: self.top.withdraw()
            except tk.TclError: pass

    def _delayed_hide(self, _):
        # Pequeno delay para permitir que o click na listbox seja processado
        self.after(150, self._maybe_hide)

    def _maybe_hide(self):
        # Só esconde se o foco não foi para a listbox
        if self.listbox and self.focus_get() == self.listbox:
            return
        self._hide_list()

    def _focus_list(self, _):
        if self.top and self.listbox and self.listbox.size() > 0:
            self.listbox.focus_set()
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)

    def _enter_key(self, _):
        if self.top and self.listbox and self.listbox.size() > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self._select_item(None)

    def _select_item(self, _):
        try:
            idx = self.listbox.curselection()[0]
            reg = self._matches[idx]
            self._ignore_next = True
            self.var.set(reg["nome"])
            self._hide_list()
            self.entry.icursor(tk.END)
            if self.on_select:
                self.on_select(reg)
        except (IndexError, AttributeError):
            pass


# ============================================================
# JANELA DE LOGIN
# ============================================================
class LoginWindow:
    def __init__(self, seguranca: Seguranca, on_success):
        self.seg = seguranca
        self.on_success = on_success

        # Altura dinâmica: primeiro acesso precisa de mais espaço
        altura = 560 if seguranca.pin_configurado() else 800

        self.root = tk.Tk()
        self.root.title("Acesso · Mercado Central BH")
        self.root.configure(bg=NAVY)
        self.root.geometry(f"460x{altura}")
        self.root.resizable(False, False)
        self._centralizar(460, altura)

        self._montar()
        self.root.mainloop()

    def _centralizar(self, w, h):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _montar(self):
        # Faixa dourada superior
        tk.Frame(self.root, bg=GOLD, height=4).pack(fill="x")

        # Logo
        try:
            from PIL import Image, ImageTk
            img = Image.open(LOGO_PATH).convert("RGBA")
            img.thumbnail((110, 110), Image.LANCZOS)
            self._logo_img = ImageTk.PhotoImage(img)
            tk.Label(self.root, image=self._logo_img, bg=NAVY).pack(pady=(30, 10))
        except Exception:
            tk.Label(self.root, text="🏛", font=("Arial", 60),
                     bg=NAVY, fg=GOLD).pack(pady=(30, 10))

        tk.Label(self.root, text="MERCADO CENTRAL",
                 font=("Georgia", 16, "bold"), bg=NAVY, fg=WHITE).pack()
        tk.Label(self.root, text="BELO HORIZONTE",
                 font=("Calibri", 9), bg=NAVY, fg=GOLD).pack(pady=(0, 20))

        tk.Frame(self.root, bg=GOLD, height=1, width=60).pack(pady=5)

        if self.seg.pin_configurado():
            self._tela_login()
        else:
            self._tela_primeiro_acesso()

        # Rodapé
        tk.Label(self.root,
                 text="Desenvolvido por Artur Gabriel Oliveira da Silva",
                 font=("Calibri", 8), bg=NAVY, fg=SLATE_LT).pack(side="bottom", pady=(0, 4))
        tk.Label(self.root, text="arturg_oliveira@outlook.com",
                 font=("Calibri", 8, "italic"), bg=NAVY, fg=GOLD_DEEP).pack(side="bottom")

    def _tela_login(self):
        tk.Label(self.root, text="Digite o PIN de acesso",
                 font=("Calibri", 12), bg=NAVY, fg=WHITE).pack(pady=(20, 10))

        self.pin_var = tk.StringVar()
        entry = tk.Entry(self.root, textvariable=self.pin_var,
                         font=("Consolas", 22), show="●",
                         justify="center", width=12,
                         relief="flat", bg=WHITE, fg=NAVY,
                         insertbackground=NAVY)
        entry.pack(ipady=10, pady=5)
        entry.focus_set()
        entry.bind("<Return>", lambda e: self._verificar())

        tk.Button(self.root, text="Entrar", command=self._verificar,
                  bg=GOLD, fg=NAVY, font=("Calibri", 12, "bold"),
                  relief="flat", cursor="hand2", width=20, pady=8,
                  activebackground=GOLD_DEEP).pack(pady=(20, 5))

        # Botão "Esqueci meu PIN" — só aparece se houver e-mail cadastrado
        if self.seg.email_recuperacao():
            tk.Button(self.root, text="Esqueci meu PIN",
                      command=self._esqueci_pin,
                      bg=NAVY, fg=GOLD_DEEP, font=("Calibri", 9, "underline"),
                      relief="flat", cursor="hand2", bd=0,
                      activebackground=NAVY, activeforeground=GOLD
                      ).pack(pady=(0, 10))

        self.status = tk.Label(self.root, text="",
                               font=("Calibri", 10), bg=NAVY, fg=DANGER)
        self.status.pack(pady=5)

    def _esqueci_pin(self):
        """Abre fluxo de recuperação de PIN por e-mail."""
        RecuperacaoWindow(self.root, self.seg,
                           on_success=self._apos_recuperacao)

    def _apos_recuperacao(self):
        """Callback: mostra mensagem de sucesso e pede para logar com novo PIN."""
        self.status.config(
            text="✓ PIN redefinido. Digite o novo PIN acima.",
            fg=SUCCESS
        )
        self.pin_var.set("")

    def _tela_primeiro_acesso(self):
        tk.Label(self.root, text="PRIMEIRO ACESSO",
                 font=("Calibri", 10, "bold"), bg=NAVY, fg=GOLD,
                 pady=0).pack(pady=(10, 2))
        tk.Label(self.root, text="Defina o PIN e o nome do operador",
                 font=("Calibri", 10), bg=NAVY, fg=WHITE).pack(pady=(0, 10))

        # Nome do operador
        tk.Label(self.root, text="Nome do operador:",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack()
        self.nome_var = tk.StringVar()
        entry_nome = tk.Entry(self.root, textvariable=self.nome_var,
                 font=("Calibri", 11), width=30, relief="flat",
                 bg=WHITE, fg=NAVY)
        entry_nome.pack(ipady=5, pady=(2, 8))
        entry_nome.focus_set()

        # PIN
        tk.Label(self.root, text="Criar PIN (4 a 8 dígitos):",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack()
        self.pin_var = tk.StringVar()
        entry_pin = tk.Entry(self.root, textvariable=self.pin_var,
                 font=("Consolas", 16), show="●", justify="center",
                 width=16, relief="flat", bg=WHITE, fg=NAVY)
        entry_pin.pack(ipady=5, pady=(2, 8))

        tk.Label(self.root, text="Confirmar PIN:",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack()
        self.pin2_var = tk.StringVar()
        entry_pin2 = tk.Entry(self.root, textvariable=self.pin2_var,
                 font=("Consolas", 16), show="●", justify="center",
                 width=16, relief="flat", bg=WHITE, fg=NAVY)
        entry_pin2.pack(ipady=5, pady=(2, 8))

        # E-mail de recuperação
        tk.Label(self.root, text="E-mail para recuperação de PIN (opcional):",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack()
        self.email_rec_var = tk.StringVar()
        entry_email = tk.Entry(self.root, textvariable=self.email_rec_var,
                 font=("Calibri", 11), width=30, relief="flat",
                 bg=WHITE, fg=NAVY)
        entry_email.pack(ipady=5, pady=(2, 12))

        # Binding Enter nos campos
        entry_nome.bind("<Return>", lambda e: entry_pin.focus_set())
        entry_pin.bind("<Return>",  lambda e: entry_pin2.focus_set())
        entry_pin2.bind("<Return>", lambda e: entry_email.focus_set())
        entry_email.bind("<Return>", lambda e: self._criar_acesso())

        # Botão grande e visível
        tk.Button(self.root, text="✓  CRIAR ACESSO", command=self._criar_acesso,
                  bg=GOLD, fg=NAVY, font=("Calibri", 12, "bold"),
                  relief="flat", cursor="hand2", width=22, pady=10,
                  activebackground=GOLD_DEEP).pack(pady=(5, 10))

        self.status = tk.Label(self.root, text="",
                               font=("Calibri", 9), bg=NAVY, fg=DANGER,
                               wraplength=400)
        self.status.pack(pady=(0, 5))

    def _verificar(self):
        if self.seg.esta_bloqueado():
            self.status.config(text=f"Bloqueado. Aguarde {self.seg.tempo_bloqueio_restante()}s.")
            return
        pin = self.pin_var.get().strip()
        if self.seg.verificar_pin(pin):
            log.info(f"Login bem-sucedido: {self.seg.operador()}")
            self.root.destroy()
            self.on_success(self.seg.operador())
        else:
            restantes = Seguranca.MAX_TENTATIVAS - self.seg.tentativas
            if self.seg.esta_bloqueado():
                self.status.config(
                    text=f"Muitas tentativas. Bloqueado por {Seguranca.BLOQUEIO_MIN} minutos.")
            else:
                self.status.config(text=f"PIN incorreto. Tentativas restantes: {restantes}")
            self.pin_var.set("")

    def _criar_acesso(self):
        nome = self.nome_var.get().strip()
        pin  = self.pin_var.get().strip()
        pin2 = self.pin2_var.get().strip()
        email_rec = self.email_rec_var.get().strip() if hasattr(self, 'email_rec_var') else ""

        if len(nome) < 3:
            self.status.config(text="Informe o nome do operador.")
            return
        if not (4 <= len(pin) <= 8) or not pin.isdigit():
            self.status.config(text="PIN deve ter entre 4 e 8 dígitos numéricos.")
            return
        if pin != pin2:
            self.status.config(text="Os PINs não conferem.")
            return
        if email_rec and not validar_email(email_rec):
            self.status.config(text="E-mail de recuperação inválido.")
            return

        self.seg.definir_pin(pin, nome, email_rec)
        aviso_email = ""
        if email_rec:
            aviso_email = f"\n\nEm caso de esquecimento, o PIN poderá ser recuperado via {email_rec}."
        messagebox.showinfo("Acesso criado",
                            f"Acesso criado para {nome}.\nGuarde o PIN em local seguro.{aviso_email}")
        self.root.destroy()
        self.on_success(nome)


# ============================================================
# JANELA DE RECUPERAÇÃO DE PIN
# ============================================================
class RecuperacaoWindow:
    """
    Fluxo de recuperação de PIN:
      1. Mostra e-mail cadastrado (parcialmente mascarado)
      2. Envia código de 6 dígitos para esse e-mail
      3. Usuário digita o código recebido
      4. Cria novo PIN
    """

    def __init__(self, parent, seguranca: Seguranca, on_success):
        self.seg = seguranca
        self.on_success = on_success
        self.etapa = 1  # 1: confirmar e-mail, 2: digitar código, 3: novo PIN

        self.win = tk.Toplevel(parent)
        self.win.title("Recuperação de PIN")
        self.win.configure(bg=NAVY)
        self.win.geometry("460x520")
        self.win.resizable(False, False)
        self.win.transient(parent)
        self.win.grab_set()

        # Centralizar sobre parent
        parent.update_idletasks()
        try:
            px = parent.winfo_rootx()
            py = parent.winfo_rooty()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            x = px + (pw - 460) // 2
            y = py + (ph - 520) // 2
            self.win.geometry(f"460x520+{x}+{y}")
        except tk.TclError:
            pass

        self._montar_base()
        self._tela_etapa1()

    def _montar_base(self):
        tk.Frame(self.win, bg=GOLD, height=4).pack(fill="x")
        tk.Label(self.win, text="RECUPERAÇÃO DE PIN",
                 font=("Calibri", 11, "bold"),
                 bg=NAVY, fg=GOLD).pack(pady=(25, 5))
        tk.Frame(self.win, bg=GOLD, height=1, width=60).pack(pady=(0, 10))

        self.frame = tk.Frame(self.win, bg=NAVY)
        self.frame.pack(fill="both", expand=True, padx=30, pady=10)

        self.status = tk.Label(self.win, text="",
                                font=("Calibri", 9), bg=NAVY, fg=DANGER,
                                wraplength=400)
        self.status.pack(pady=(0, 10))

    def _limpar_frame(self):
        for w in self.frame.winfo_children():
            w.destroy()
        self.status.config(text="")

    @staticmethod
    def _mascarar_email(email: str) -> str:
        """joao@email.com → j***@email.com"""
        if "@" not in email:
            return email
        local, dom = email.split("@", 1)
        if len(local) <= 2:
            return f"{local[0]}***@{dom}"
        return f"{local[0]}{'*' * (len(local) - 2)}{local[-1]}@{dom}"

    # ---------- ETAPA 1: Confirmar envio para e-mail cadastrado ----------
    def _tela_etapa1(self):
        self._limpar_frame()
        email = self.seg.email_recuperacao()
        if not email:
            tk.Label(self.frame,
                     text="Nenhum e-mail de recuperação foi cadastrado.\n\n"
                          "Para redefinir o PIN, feche o programa, apague o arquivo "
                          "config.json na pasta do aplicativo e reabra para criar um "
                          "novo acesso.",
                     font=("Calibri", 10), bg=NAVY, fg=WHITE,
                     wraplength=380, justify="left").pack(pady=20)
            tk.Button(self.frame, text="Entendi",
                      command=self.win.destroy,
                      bg=GOLD, fg=NAVY, font=("Calibri", 10, "bold"),
                      relief="flat", cursor="hand2", width=18, pady=6
                      ).pack(pady=10)
            return

        tk.Label(self.frame,
                 text="Um código de 6 dígitos será enviado para o e-mail cadastrado:",
                 font=("Calibri", 10), bg=NAVY, fg=WHITE,
                 wraplength=380, justify="left").pack(pady=(10, 8))

        tk.Label(self.frame, text=self._mascarar_email(email),
                 font=("Calibri", 13, "bold"),
                 bg=NAVY, fg=GOLD).pack(pady=5)

        tk.Label(self.frame,
                 text="O código é válido por 15 minutos.",
                 font=("Calibri", 9, "italic"),
                 bg=NAVY, fg=SLATE_LT).pack(pady=(8, 20))

        tk.Button(self.frame, text="📧  Enviar código",
                  command=self._enviar_codigo,
                  bg=GOLD, fg=NAVY, font=("Calibri", 11, "bold"),
                  relief="flat", cursor="hand2", width=25, pady=8,
                  activebackground=GOLD_DEEP).pack(pady=5)

        tk.Button(self.frame, text="Cancelar",
                  command=self.win.destroy,
                  bg=NAVY, fg=SLATE_LT, font=("Calibri", 9),
                  relief="flat", cursor="hand2", bd=0
                  ).pack(pady=(15, 5))

    def _enviar_codigo(self):
        email = self.seg.email_recuperacao()
        if not SMTP_USER or not SMTP_PASS:
            messagebox.showerror(
                "SMTP não configurado",
                "As credenciais de e-mail não estão configuradas no .env.\n\n"
                "Sem o .env configurado, não é possível enviar o código "
                "de recuperação.",
                parent=self.win
            )
            return

        self.status.config(text="Enviando código...", fg=SLATE_LT)
        self.win.update()

        codigo = self.seg.gerar_token_recuperacao()
        assunto = "Código de recuperação de PIN — Mercado Central BH"
        corpo = (
            f"Olá,\n\n"
            f"Você solicitou a recuperação do PIN do Emissor de Certificados "
            f"do Mercado Central de Belo Horizonte.\n\n"
            f"Seu código de recuperação é:\n\n"
            f"    {codigo}\n\n"
            f"Este código expira em 15 minutos.\n\n"
            f"Se você não solicitou a recuperação, ignore este e-mail.\n\n"
            f"Atenciosamente,\n"
            f"Sistema Emissor de Certificados\n"
            f"Mercado Central de Belo Horizonte\n"
        )

        ok, msg = Envio.email_simples(email, assunto, corpo)
        if ok:
            log.info(f"Código de recuperação enviado para {self._mascarar_email(email)}")
            self._tela_etapa2()
        else:
            self.status.config(text=f"Falha ao enviar: {msg}", fg=DANGER)

    # ---------- ETAPA 2: Digitar código ----------
    def _tela_etapa2(self):
        self._limpar_frame()

        tk.Label(self.frame,
                 text="Código enviado! Digite o código de 6 dígitos que você recebeu:",
                 font=("Calibri", 10), bg=NAVY, fg=WHITE,
                 wraplength=380, justify="left").pack(pady=(10, 15))

        self.codigo_var = tk.StringVar()
        entry = tk.Entry(self.frame, textvariable=self.codigo_var,
                          font=("Consolas", 22), justify="center",
                          width=10, relief="flat", bg=WHITE, fg=NAVY,
                          insertbackground=NAVY)
        entry.pack(ipady=8, pady=5)
        entry.focus_set()
        entry.bind("<Return>", lambda e: self._validar_codigo())

        tk.Label(self.frame,
                 text="Verifique também a caixa de spam.",
                 font=("Calibri", 9, "italic"),
                 bg=NAVY, fg=SLATE_LT).pack(pady=(8, 15))

        tk.Button(self.frame, text="Validar código",
                  command=self._validar_codigo,
                  bg=GOLD, fg=NAVY, font=("Calibri", 11, "bold"),
                  relief="flat", cursor="hand2", width=25, pady=8,
                  activebackground=GOLD_DEEP).pack(pady=5)

        tk.Button(self.frame, text="« Reenviar código",
                  command=self._enviar_codigo,
                  bg=NAVY, fg=SLATE_LT, font=("Calibri", 9, "underline"),
                  relief="flat", cursor="hand2", bd=0
                  ).pack(pady=(15, 5))

    def _validar_codigo(self):
        codigo = self.codigo_var.get().strip()
        if not codigo or not codigo.isdigit() or len(codigo) != 6:
            self.status.config(text="Digite os 6 dígitos do código.", fg=DANGER)
            return
        if self.seg.validar_token_recuperacao(codigo):
            self._tela_etapa3()
        else:
            self.status.config(
                text="Código inválido ou expirado. Verifique e tente novamente.",
                fg=DANGER
            )

    # ---------- ETAPA 3: Criar novo PIN ----------
    def _tela_etapa3(self):
        self._limpar_frame()

        tk.Label(self.frame, text="✓  Código validado!",
                 font=("Calibri", 11, "bold"), bg=NAVY, fg=SUCCESS).pack(pady=(10, 5))
        tk.Label(self.frame, text="Crie seu novo PIN de acesso:",
                 font=("Calibri", 10), bg=NAVY, fg=WHITE).pack(pady=(0, 15))

        tk.Label(self.frame, text="Novo PIN (4 a 8 dígitos):",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack()
        self.novo_pin_var = tk.StringVar()
        entry_pin = tk.Entry(self.frame, textvariable=self.novo_pin_var,
                              font=("Consolas", 16), show="●", justify="center",
                              width=16, relief="flat", bg=WHITE, fg=NAVY)
        entry_pin.pack(ipady=5, pady=(2, 10))
        entry_pin.focus_set()

        tk.Label(self.frame, text="Confirmar novo PIN:",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack()
        self.novo_pin2_var = tk.StringVar()
        entry_pin2 = tk.Entry(self.frame, textvariable=self.novo_pin2_var,
                               font=("Consolas", 16), show="●", justify="center",
                               width=16, relief="flat", bg=WHITE, fg=NAVY)
        entry_pin2.pack(ipady=5, pady=(2, 15))

        entry_pin.bind("<Return>", lambda e: entry_pin2.focus_set())
        entry_pin2.bind("<Return>", lambda e: self._definir_novo_pin())

        tk.Button(self.frame, text="✓  Confirmar novo PIN",
                  command=self._definir_novo_pin,
                  bg=GOLD, fg=NAVY, font=("Calibri", 11, "bold"),
                  relief="flat", cursor="hand2", width=25, pady=8,
                  activebackground=GOLD_DEEP).pack(pady=5)

    def _definir_novo_pin(self):
        pin = self.novo_pin_var.get().strip()
        pin2 = self.novo_pin2_var.get().strip()

        if not (4 <= len(pin) <= 8) or not pin.isdigit():
            self.status.config(text="PIN deve ter entre 4 e 8 dígitos numéricos.",
                                fg=DANGER)
            return
        if pin != pin2:
            self.status.config(text="Os PINs não conferem.", fg=DANGER)
            return

        self.seg.redefinir_pin_apos_recuperacao(pin)
        messagebox.showinfo("PIN redefinido",
                             "Seu PIN foi redefinido com sucesso!\n"
                             "Use o novo PIN na próxima tela de login.",
                             parent=self.win)
        self.win.destroy()
        if self.on_success:
            self.on_success()


# ============================================================
# JANELA DE CONFIGURAÇÕES
# ============================================================
class ConfigWindow:
    """
    Tela para editar dados que aparecem no certificado:
    - Nome e cargo do signatário
    - Imagem de assinatura digital
    - Data padrão da Assembleia
    - Nome do operador
    """

    def __init__(self, parent, seguranca: Seguranca, on_close=None):
        self.seg = seguranca
        self.on_close = on_close
        self.assinatura_path_var = tk.StringVar(
            value=seguranca.signatario_assinatura_path() or ""
        )

        self.win = tk.Toplevel(parent)
        self.win.title("Configurações do Certificado")
        self.win.configure(bg=IVORY)
        # Usar tamanho menor agora que tem scroll — cabe em qualquer tela
        self.win.geometry("640x720")
        self.win.minsize(620, 500)
        self.win.transient(parent)
        self.win.grab_set()

        # Centralizar sobre parent
        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        x = px + (pw - 640) // 2
        y = max(10, py + (ph - 720) // 2)
        # Se for passar do fundo da tela, limitar
        screen_h = parent.winfo_screenheight()
        if y + 720 > screen_h - 40:
            y = max(10, screen_h - 760)
        self.win.geometry(f"640x720+{x}+{y}")

        self._montar()

    def _montar(self):
        # Header
        header = tk.Frame(self.win, bg=NAVY, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="⚙  Configurações do Certificado",
                 font=("Georgia", 14, "bold"),
                 bg=NAVY, fg=WHITE).pack(side="left", padx=20, pady=15)
        tk.Frame(self.win, bg=GOLD, height=2).pack(fill="x")

        # ----- FOOTER FIXO NA BASE (packed ANTES do body) -----
        footer = tk.Frame(self.win, bg=IVORY, pady=12)
        footer.pack(side="bottom", fill="x")
        tk.Frame(footer, bg=DIVIDER, height=1).pack(side="top", fill="x", pady=(0, 12))

        tk.Button(footer, text="Cancelar", command=self._cancelar,
                  bg=IVORY, fg=SLATE, font=("Calibri", 10),
                  relief="solid", bd=1, cursor="hand2",
                  padx=20, pady=8
                  ).pack(side="right", padx=(8, 20))

        tk.Button(footer, text="💾  Salvar alterações", command=self._salvar,
                  bg=GOLD, fg=NAVY, font=("Calibri", 11, "bold"),
                  relief="flat", cursor="hand2", padx=20, pady=8,
                  activebackground=GOLD_DEEP
                  ).pack(side="right")

        tk.Button(footer, text="📝  Abrir modelo no Word",
                  command=self._abrir_modelo,
                  bg=NAVY, fg=WHITE, font=("Calibri", 10, "bold"),
                  relief="flat", cursor="hand2", padx=14, pady=8,
                  activebackground=NAVY_DARK, activeforeground=WHITE
                  ).pack(side="left", padx=(20, 0))

        # ----- CORPO COM SCROLL -----
        # Container que segura o Canvas + Scrollbar
        body_container = tk.Frame(self.win, bg=IVORY)
        body_container.pack(fill="both", expand=True)

        # Canvas que permite scroll vertical
        canvas = tk.Canvas(body_container, bg=IVORY, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(body_container, orient="vertical",
                                    command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Frame interno onde o conteúdo de fato vai (scroll atua nele)
        body = tk.Frame(canvas, bg=IVORY, padx=25, pady=20)
        canvas_window = canvas.create_window((0, 0), window=body, anchor="nw")

        # Atualizar scroll region quando conteúdo muda
        def _atualizar_scroll(_=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Ajustar largura do frame interno ao canvas
            canvas_w = canvas.winfo_width()
            canvas.itemconfig(canvas_window, width=canvas_w)

        body.bind("<Configure>", _atualizar_scroll)
        canvas.bind("<Configure>", _atualizar_scroll)

        # Permitir scroll pela rodinha do mouse
        def _on_mousewheel(event):
            # Windows: event.delta é ±120; Linux: event.num é 4/5
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

        def _bind_scroll(_):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_scroll(_):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        # Só ativa o scroll da rodinha quando o mouse estiver sobre a janela
        self.win.bind("<Enter>", _bind_scroll)
        self.win.bind("<Leave>", _unbind_scroll)

        # ---------- SIGNATÁRIO ----------
        frame_sig = tk.LabelFrame(body, text="  Signatário do certificado  ",
                                   font=("Calibri", 10, "bold"),
                                   bg=IVORY, fg=NAVY, bd=1, relief="solid",
                                   labelanchor="nw", padx=15, pady=12)
        frame_sig.pack(fill="x", pady=(0, 12))

        tk.Label(frame_sig, text="Nome completo do signatário:",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(anchor="w")
        self.v_sig_nome = tk.StringVar(value=self.seg.signatario_nome())
        tk.Entry(frame_sig, textvariable=self.v_sig_nome,
                 font=("Calibri", 11), relief="solid", bd=1, bg=WHITE,
                 highlightthickness=1, highlightbackground=DIVIDER,
                 highlightcolor=GOLD
                 ).pack(fill="x", ipady=5, pady=(3, 10))

        tk.Label(frame_sig, text="Cargo:",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(anchor="w")
        self.v_sig_cargo = tk.StringVar(value=self.seg.signatario_cargo())
        tk.Entry(frame_sig, textvariable=self.v_sig_cargo,
                 font=("Calibri", 11), relief="solid", bd=1, bg=WHITE,
                 highlightthickness=1, highlightbackground=DIVIDER,
                 highlightcolor=GOLD
                 ).pack(fill="x", ipady=5, pady=(3, 0))

        # ---------- ASSINATURA DIGITAL ----------
        frame_ass = tk.LabelFrame(body, text="  Imagem da assinatura (opcional)  ",
                                   font=("Calibri", 10, "bold"),
                                   bg=IVORY, fg=NAVY, bd=1, relief="solid",
                                   labelanchor="nw", padx=15, pady=12)
        frame_ass.pack(fill="x", pady=(0, 12))

        tk.Label(frame_ass,
                 text="Use uma imagem PNG com fundo transparente (rubrica escaneada).",
                 font=("Calibri", 9, "italic"),
                 bg=IVORY, fg=SLATE, wraplength=520, justify="left"
                 ).pack(anchor="w", pady=(0, 8))

        btns_ass = tk.Frame(frame_ass, bg=IVORY)
        btns_ass.pack(fill="x")

        tk.Button(btns_ass, text="📎  Escolher imagem...",
                  command=self._escolher_assinatura,
                  bg=NAVY, fg=WHITE, font=("Calibri", 10),
                  relief="flat", cursor="hand2", padx=12, pady=6,
                  activebackground=NAVY_DARK, activeforeground=WHITE
                  ).pack(side="left", padx=(0, 8))

        tk.Button(btns_ass, text="Remover",
                  command=self._remover_assinatura,
                  bg=IVORY, fg=DANGER, font=("Calibri", 9),
                  relief="solid", bd=1, cursor="hand2", padx=10, pady=5
                  ).pack(side="left")

        # Preview
        self.lbl_preview = tk.Label(frame_ass, text="",
                                     font=("Calibri", 9, "italic"),
                                     bg=IVORY, fg=SLATE_LT, anchor="w",
                                     wraplength=520, justify="left")
        self.lbl_preview.pack(fill="x", pady=(8, 0))
        self._atualizar_preview()

        # ---------- DATA PADRÃO ----------
        frame_data = tk.LabelFrame(body, text="  Dados padrão  ",
                                    font=("Calibri", 10, "bold"),
                                    bg=IVORY, fg=NAVY, bd=1, relief="solid",
                                    labelanchor="nw", padx=15, pady=12)
        frame_data.pack(fill="x", pady=(0, 12))

        tk.Label(frame_data, text="Data da Assembleia (sugestão inicial):",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(anchor="w")
        self.v_data = tk.StringVar(value=self.seg.data_assembleia_padrao())
        tk.Entry(frame_data, textvariable=self.v_data,
                 font=("Calibri", 11), relief="solid", bd=1, bg=WHITE,
                 highlightthickness=1, highlightbackground=DIVIDER,
                 highlightcolor=GOLD
                 ).pack(fill="x", ipady=5, pady=(3, 0))

        # ---------- SEGURANÇA ----------
        frame_seg = tk.LabelFrame(body, text="  Segurança  ",
                                    font=("Calibri", 10, "bold"),
                                    bg=IVORY, fg=NAVY, bd=1, relief="solid",
                                    labelanchor="nw", padx=15, pady=12)
        frame_seg.pack(fill="x", pady=(0, 12))

        # E-mail de recuperação
        tk.Label(frame_seg, text="E-mail de recuperação (para \"Esqueci meu PIN\"):",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(anchor="w")
        self.v_email_rec = tk.StringVar(value=self.seg.email_recuperacao())
        tk.Entry(frame_seg, textvariable=self.v_email_rec,
                 font=("Calibri", 11), relief="solid", bd=1, bg=WHITE,
                 highlightthickness=1, highlightbackground=DIVIDER,
                 highlightcolor=GOLD
                 ).pack(fill="x", ipady=5, pady=(3, 10))

        # Linha: botão trocar PIN
        tk.Label(frame_seg, text="Trocar PIN de acesso:",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(anchor="w", pady=(5, 3))

        tk.Button(frame_seg, text="🔑  Trocar PIN agora...",
                  command=self._abrir_trocar_pin,
                  bg=NAVY, fg=WHITE, font=("Calibri", 10),
                  relief="flat", cursor="hand2", padx=12, pady=6,
                  activebackground=NAVY_DARK, activeforeground=WHITE
                  ).pack(anchor="w")

        # ---------- DICA ----------
        dica = tk.Label(body,
                        text="💡  Para personalizar o texto, cores ou layout do certificado, "
                             "clique em \"📝 Abrir modelo no Word\" abaixo. Preserve os "
                             "placeholders {{ nome }}, {{ matricula }}, {{ qr_code }}, etc.",
                        font=("Calibri", 9, "italic"),
                        bg=IVORY, fg=SLATE_LT, wraplength=540, justify="left",
                        anchor="w")
        dica.pack(fill="x", pady=(0, 10))

    # ---------- AÇÕES ----------
    def _escolher_assinatura(self):
        path = filedialog.askopenfilename(
            title="Selecione a imagem da assinatura",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg"), ("Todos", "*.*")],
            parent=self.win
        )
        if path:
            # Copia para pasta do programa para evitar caminho quebrado
            try:
                import shutil
                destino = BASE_DIR / f"assinatura{Path(path).suffix.lower()}"
                shutil.copy2(path, destino)
                self.assinatura_path_var.set(str(destino))
                self._atualizar_preview()
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível copiar a imagem:\n{e}",
                                     parent=self.win)

    def _remover_assinatura(self):
        self.assinatura_path_var.set("")
        self._atualizar_preview()

    def _atualizar_preview(self):
        path = self.assinatura_path_var.get()
        if path and Path(path).exists():
            try:
                tam = Path(path).stat().st_size / 1024
                self.lbl_preview.config(
                    text=f"✓ Arquivo carregado: {Path(path).name}  ({tam:.1f} KB)",
                    fg=SUCCESS
                )
            except Exception:
                self.lbl_preview.config(text=f"✓ {Path(path).name}", fg=SUCCESS)
        else:
            self.lbl_preview.config(
                text="Nenhuma assinatura carregada. O certificado usará linha em branco.",
                fg=SLATE_LT
            )

    def _abrir_modelo(self):
        if not MODELO_PATH.exists():
            messagebox.showerror("Erro", f"Modelo não encontrado: {MODELO_PATH.name}",
                                 parent=self.win)
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(MODELO_PATH))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(MODELO_PATH)])
            else:
                subprocess.Popen(["xdg-open", str(MODELO_PATH)])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o modelo:\n{e}",
                                 parent=self.win)

    def _salvar(self):
        nome = self.v_sig_nome.get().strip()
        cargo = self.v_sig_cargo.get().strip()
        data = self.v_data.get().strip()
        email_rec = self.v_email_rec.get().strip()
        assinatura = self.assinatura_path_var.get().strip() or None

        if len(nome) < 3:
            messagebox.showwarning("Atenção", "Informe o nome do signatário.",
                                    parent=self.win)
            return
        if len(cargo) < 2:
            messagebox.showwarning("Atenção", "Informe o cargo do signatário.",
                                    parent=self.win)
            return
        if len(data) < 5:
            messagebox.showwarning("Atenção", "Informe a data padrão da Assembleia.",
                                    parent=self.win)
            return
        if email_rec and not validar_email(email_rec):
            messagebox.showwarning("Atenção",
                                    "E-mail de recuperação inválido.\n"
                                    "Deixe em branco ou informe um e-mail válido.",
                                    parent=self.win)
            return

        self.seg.atualizar_signatario(nome, cargo, assinatura, data)
        self.seg.definir_email_recuperacao(email_rec)
        messagebox.showinfo("Salvo",
                             "Configurações salvas.\nOs próximos certificados gerados "
                             "já usarão os novos dados.",
                             parent=self.win)
        self.win.destroy()
        if self.on_close:
            self.on_close()

    def _abrir_trocar_pin(self):
        """Abre janela para trocar o PIN atual."""
        TrocarPinWindow(self.win, self.seg)

    def _cancelar(self):
        self.win.destroy()


# ============================================================
# JANELA DE TROCA DE PIN
# ============================================================
class TrocarPinWindow:
    """Janela modal para trocar o PIN atual (requer PIN atual + novo duas vezes)."""

    def __init__(self, parent, seguranca: Seguranca):
        self.seg = seguranca

        self.win = tk.Toplevel(parent)
        self.win.title("Trocar PIN")
        self.win.configure(bg=NAVY)
        self.win.geometry("420x460")
        self.win.resizable(False, False)
        self.win.transient(parent)
        self.win.grab_set()

        # Centralizar
        parent.update_idletasks()
        try:
            px = parent.winfo_rootx(); py = parent.winfo_rooty()
            pw = parent.winfo_width();  ph = parent.winfo_height()
            x = px + (pw - 420) // 2
            y = py + (ph - 460) // 2
            self.win.geometry(f"420x460+{x}+{y}")
        except tk.TclError:
            pass

        self._montar()

    def _montar(self):
        tk.Frame(self.win, bg=GOLD, height=4).pack(fill="x")
        tk.Label(self.win, text="🔑  TROCAR PIN",
                 font=("Calibri", 12, "bold"),
                 bg=NAVY, fg=GOLD).pack(pady=(25, 5))
        tk.Frame(self.win, bg=GOLD, height=1, width=60).pack(pady=(0, 15))

        frame = tk.Frame(self.win, bg=NAVY)
        frame.pack(fill="both", expand=True, padx=30)

        # PIN atual
        tk.Label(frame, text="PIN atual:",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack(anchor="w")
        self.atual_var = tk.StringVar()
        e1 = tk.Entry(frame, textvariable=self.atual_var,
                       font=("Consolas", 16), show="●", justify="center",
                       width=18, relief="flat", bg=WHITE, fg=NAVY)
        e1.pack(ipady=5, pady=(3, 12))
        e1.focus_set()

        # PIN novo
        tk.Label(frame, text="Novo PIN (4 a 8 dígitos):",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack(anchor="w")
        self.novo_var = tk.StringVar()
        e2 = tk.Entry(frame, textvariable=self.novo_var,
                       font=("Consolas", 16), show="●", justify="center",
                       width=18, relief="flat", bg=WHITE, fg=NAVY)
        e2.pack(ipady=5, pady=(3, 12))

        # Confirmar
        tk.Label(frame, text="Confirmar novo PIN:",
                 font=("Calibri", 9), bg=NAVY, fg=SLATE_LT).pack(anchor="w")
        self.novo2_var = tk.StringVar()
        e3 = tk.Entry(frame, textvariable=self.novo2_var,
                       font=("Consolas", 16), show="●", justify="center",
                       width=18, relief="flat", bg=WHITE, fg=NAVY)
        e3.pack(ipady=5, pady=(3, 18))

        e1.bind("<Return>", lambda e: e2.focus_set())
        e2.bind("<Return>", lambda e: e3.focus_set())
        e3.bind("<Return>", lambda e: self._salvar())

        tk.Button(frame, text="✓  Confirmar troca",
                  command=self._salvar,
                  bg=GOLD, fg=NAVY, font=("Calibri", 11, "bold"),
                  relief="flat", cursor="hand2", width=22, pady=8,
                  activebackground=GOLD_DEEP).pack(pady=5)

        tk.Button(frame, text="Cancelar",
                  command=self.win.destroy,
                  bg=NAVY, fg=SLATE_LT, font=("Calibri", 9),
                  relief="flat", cursor="hand2", bd=0
                  ).pack(pady=(8, 5))

        self.status = tk.Label(self.win, text="",
                                font=("Calibri", 9), bg=NAVY, fg=DANGER,
                                wraplength=380)
        self.status.pack(pady=(0, 10))

    def _salvar(self):
        atual = self.atual_var.get().strip()
        novo = self.novo_var.get().strip()
        novo2 = self.novo2_var.get().strip()

        if not atual:
            self.status.config(text="Informe o PIN atual.")
            return
        if novo != novo2:
            self.status.config(text="O novo PIN e a confirmação não conferem.")
            return

        ok, msg = self.seg.trocar_pin(atual, novo)
        if ok:
            messagebox.showinfo("PIN alterado", msg, parent=self.win)
            self.win.destroy()
        else:
            self.status.config(text=msg)
            self.atual_var.set("")


# ============================================================
# JANELA DE HISTÓRICO DE EMISSÕES
# ============================================================
class HistoricoWindow:
    """
    Tabela com todas as emissões feitas, com filtros (texto/data).
    Permite: abrir PDF, reimprimir (novo código), reenviar por email/WhatsApp.
    """

    def __init__(self, parent, seguranca: Seguranca,
                 historico: "HistoricoEmissoes",
                 gerador: "GeradorCertificado",
                 operador: str):
        self.seg = seguranca
        self.historico = historico
        self.gerador = gerador
        self.operador = operador

        self.win = tk.Toplevel(parent)
        self.win.title("Histórico de Emissões")
        self.win.configure(bg=IVORY)
        self.win.geometry("1180x650")
        self.win.transient(parent)

        # Centralizar
        parent.update_idletasks()
        try:
            px = parent.winfo_rootx(); py = parent.winfo_rooty()
            pw = parent.winfo_width();  ph = parent.winfo_height()
            x = max(10, px + (pw - 1180) // 2)
            y = max(20, py + (ph - 650) // 2)
            self.win.geometry(f"1180x650+{x}+{y}")
        except tk.TclError:
            pass

        self._montar()
        self._recarregar()

    def _montar(self):
        # ---------- HEADER ----------
        header = tk.Frame(self.win, bg=NAVY, height=55)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="📋  Histórico de Emissões",
                 font=("Georgia", 14, "bold"),
                 bg=NAVY, fg=WHITE).pack(side="left", padx=20, pady=14)
        tk.Button(header, text="✕  Fechar",
                  command=self.win.destroy,
                  bg=NAVY, fg=SLATE_LT, font=("Calibri", 10),
                  relief="flat", cursor="hand2", bd=0,
                  activebackground=NAVY
                  ).pack(side="right", padx=15, pady=14)
        tk.Frame(self.win, bg=GOLD, height=2).pack(fill="x")

        # ---------- FILTROS ----------
        filtros = tk.LabelFrame(self.win, text="  Filtros  ",
                                 font=("Calibri", 10, "bold"),
                                 bg=IVORY, fg=NAVY, bd=1, relief="solid",
                                 labelanchor="nw", padx=15, pady=10)
        filtros.pack(fill="x", padx=15, pady=(10, 8))

        # Linha 1: busca por texto
        linha1 = tk.Frame(filtros, bg=IVORY)
        linha1.pack(fill="x")
        tk.Label(linha1, text="Buscar (nome/matrícula/código):",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(side="left", padx=(0, 6))
        self.v_filtro = tk.StringVar()
        e_filtro = tk.Entry(linha1, textvariable=self.v_filtro,
                             font=("Calibri", 10), relief="solid", bd=1, bg=WHITE,
                             highlightthickness=1, highlightbackground=DIVIDER,
                             highlightcolor=GOLD, width=30)
        e_filtro.pack(side="left", ipady=3, padx=(0, 12))
        e_filtro.bind("<KeyRelease>", lambda e: self._recarregar())

        tk.Label(linha1, text="De:",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(side="left", padx=(5, 4))
        self.v_data_de = tk.StringVar()
        e_de = tk.Entry(linha1, textvariable=self.v_data_de,
                         font=("Calibri", 10), relief="solid", bd=1, bg=WHITE,
                         highlightthickness=1, highlightbackground=DIVIDER,
                         highlightcolor=GOLD, width=12)
        e_de.pack(side="left", ipady=3, padx=(0, 8))
        e_de.bind("<KeyRelease>", lambda e: self._recarregar())

        tk.Label(linha1, text="Até:",
                 font=("Calibri", 9, "bold"), bg=IVORY, fg=NAVY
                 ).pack(side="left", padx=(5, 4))
        self.v_data_ate = tk.StringVar()
        e_ate = tk.Entry(linha1, textvariable=self.v_data_ate,
                          font=("Calibri", 10), relief="solid", bd=1, bg=WHITE,
                          highlightthickness=1, highlightbackground=DIVIDER,
                          highlightcolor=GOLD, width=12)
        e_ate.pack(side="left", ipady=3, padx=(0, 8))
        e_ate.bind("<KeyRelease>", lambda e: self._recarregar())

        tk.Label(linha1, text="(formato: AAAA-MM-DD)",
                 font=("Calibri", 8, "italic"), bg=IVORY, fg=SLATE_LT
                 ).pack(side="left", padx=(4, 8))

        tk.Button(linha1, text="Limpar",
                  command=self._limpar_filtros,
                  bg=IVORY, fg=SLATE, font=("Calibri", 9),
                  relief="solid", bd=1, cursor="hand2", padx=8, pady=3
                  ).pack(side="right")

        # ---------- TABELA ----------
        table_frame = tk.Frame(self.win, bg=IVORY)
        table_frame.pack(fill="both", expand=True, padx=15, pady=5)

        # Scrollbars
        scroll_y = tk.Scrollbar(table_frame, orient="vertical")
        scroll_y.pack(side="right", fill="y")

        cols = ("timestamp", "codigo", "matricula", "nome",
                "email", "email_env", "zap_env", "operador")
        self.tree = ttk.Treeview(
            table_frame, columns=cols, show="headings",
            yscrollcommand=scroll_y.set, height=15
        )
        scroll_y.config(command=self.tree.yview)

        # Estilo
        style = ttk.Style()
        style.configure("Treeview",
                        rowheight=26, font=("Calibri", 10),
                        background=WHITE, fieldbackground=WHITE)
        style.configure("Treeview.Heading",
                        font=("Calibri", 10, "bold"),
                        background=NAVY, foreground=WHITE)
        style.map("Treeview",
                  background=[("selected", GOLD)],
                  foreground=[("selected", NAVY)])

        # Colunas
        self.tree.heading("timestamp", text="Data / Hora")
        self.tree.heading("codigo", text="Código")
        self.tree.heading("matricula", text="Matrícula")
        self.tree.heading("nome", text="Nome do Associado")
        self.tree.heading("email", text="E-mail")
        self.tree.heading("email_env", text="E-mail enviado")
        self.tree.heading("zap_env", text="WhatsApp enviado")
        self.tree.heading("operador", text="Operador")

        self.tree.column("timestamp", width=135, anchor="w", minwidth=120)
        self.tree.column("codigo", width=135, anchor="w", minwidth=120)
        self.tree.column("matricula", width=80, anchor="center", minwidth=70)
        self.tree.column("nome", width=240, anchor="w", minwidth=180)
        self.tree.column("email", width=195, anchor="w", minwidth=150)
        self.tree.column("email_env", width=105, anchor="center", minwidth=95)
        self.tree.column("zap_env", width=125, anchor="center", minwidth=115)
        self.tree.column("operador", width=125, anchor="w", minwidth=100)

        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-Button-1>", lambda e: self._abrir_pdf())

        # ---------- STATUS ----------
        self.lbl_total = tk.Label(self.win, text="",
                                    font=("Calibri", 9, "italic"),
                                    bg=IVORY, fg=SLATE)
        self.lbl_total.pack(pady=(2, 0))

        # ---------- BOTÕES DE AÇÃO ----------
        botoes = tk.Frame(self.win, bg=IVORY)
        botoes.pack(fill="x", padx=15, pady=10)

        tk.Button(botoes, text="👁  Abrir PDF",
                  command=self._abrir_pdf,
                  bg=NAVY, fg=WHITE, font=("Calibri", 10, "bold"),
                  relief="flat", cursor="hand2", padx=15, pady=7,
                  activebackground=NAVY_DARK, activeforeground=WHITE
                  ).pack(side="left", padx=(0, 6))

        tk.Button(botoes, text="🔄  Reimprimir (novo código)",
                  command=self._reimprimir,
                  bg=GOLD, fg=NAVY, font=("Calibri", 10, "bold"),
                  relief="flat", cursor="hand2", padx=15, pady=7,
                  activebackground=GOLD_DEEP
                  ).pack(side="left", padx=(0, 6))

        tk.Button(botoes, text="📧  Reenviar E-mail",
                  command=self._reenviar_email,
                  bg=SUCCESS, fg=WHITE, font=("Calibri", 10, "bold"),
                  relief="flat", cursor="hand2", padx=15, pady=7
                  ).pack(side="left", padx=(0, 6))

        tk.Button(botoes, text="📱  Reenviar WhatsApp",
                  command=self._reenviar_whatsapp,
                  bg="#25D366", fg=WHITE, font=("Calibri", 10, "bold"),
                  relief="flat", cursor="hand2", padx=15, pady=7
                  ).pack(side="left", padx=(0, 6))

        tk.Button(botoes, text="📁  Abrir pasta",
                  command=self._abrir_pasta,
                  bg=IVORY, fg=NAVY, font=("Calibri", 9),
                  relief="solid", bd=1, cursor="hand2", padx=10, pady=7
                  ).pack(side="right")

    # ---------- AÇÕES ----------
    def _recarregar(self):
        # Limpar tabela
        for item in self.tree.get_children():
            self.tree.delete(item)

        registros = self.historico.listar(
            filtro_texto=self.v_filtro.get(),
            data_de=self.v_data_de.get().strip(),
            data_ate=self.v_data_ate.get().strip()
        )

        for r in registros:
            ts = r.get("timestamp", "")[:16].replace("T", " ")
            email_env = "✓ Sim" if r.get("enviado_email") == "S" else "—"
            zap_env = "✓ Sim" if r.get("enviado_whatsapp") == "S" else "—"
            self.tree.insert("", tk.END, values=(
                ts,
                r.get("codigo", ""),
                r.get("matricula", ""),
                r.get("nome", ""),
                r.get("email", ""),
                email_env,
                zap_env,
                r.get("operador", "")
            ), tags=(r.get("codigo", ""),))

        total = len(registros)
        self.lbl_total.config(text=f"Exibindo {total} registro(s)" if total
                              else "Nenhum registro encontrado com os filtros atuais.")

    def _limpar_filtros(self):
        self.v_filtro.set("")
        self.v_data_de.set("")
        self.v_data_ate.set("")
        self._recarregar()

    def _selecionado(self) -> dict | None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selecione",
                                  "Selecione um certificado na lista primeiro.",
                                  parent=self.win)
            return None
        item = self.tree.item(sel[0])
        codigo = item["values"][1]
        return self.historico.buscar_por_codigo(str(codigo))

    def _abrir_pdf(self):
        reg = self._selecionado()
        if not reg:
            return
        arquivo = reg.get("arquivo_pdf", "")
        if not arquivo or not Path(arquivo).exists():
            messagebox.showwarning("Arquivo não encontrado",
                                    "O PDF original não foi encontrado na pasta.\n"
                                    "Use 'Reimprimir' para gerar um novo certificado.",
                                    parent=self.win)
            return
        try:
            if sys.platform == "win32":
                os.startfile(arquivo)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", arquivo])
            else:
                subprocess.Popen(["xdg-open", arquivo])
        except Exception as e:
            messagebox.showerror("Erro", str(e), parent=self.win)

    def _reimprimir(self):
        reg = self._selecionado()
        if not reg:
            return
        nome = reg.get("nome", "")
        matr = reg.get("matricula", "")
        data_assem = reg.get("data_assembleia", "") or self.seg.data_assembleia_padrao()

        if not messagebox.askyesno(
            "Reimprimir certificado",
            f"Gerar novo certificado para:\n\n"
            f"{nome}\nMatrícula: {matr}\n\n"
            f"Um novo código de autenticidade será criado.\n"
            f"O certificado original fica preservado.",
            parent=self.win
        ):
            return

        arquivo, codigo, erro = self.gerador.gerar(
            nome, matr, data_assem,
            signatario_nome=self.seg.signatario_nome(),
            signatario_cargo=self.seg.signatario_cargo(),
            assinatura_path=self.seg.signatario_assinatura_path(),
        )
        if not arquivo:
            messagebox.showerror("Erro", erro or "Falha na geração.",
                                   parent=self.win)
            return

        # Registrar no histórico
        self.historico.registrar({
            "codigo": codigo,
            "matricula": matr,
            "nome": nome,
            "email": reg.get("email", ""),
            "telefone": reg.get("telefone", ""),
            "data_assembleia": data_assem,
            "arquivo_pdf": str(arquivo),
            "operador": self.operador,
            "enviado_email": "",
            "enviado_whatsapp": "",
        })
        registrar_auditoria(self.operador, matr, nome, "", "reimpressao",
                             "SUCESSO", codigo)
        messagebox.showinfo("Reimpressão concluída",
                             f"Novo certificado gerado!\n\n"
                             f"Código: {codigo}",
                             parent=self.win)
        self._recarregar()
        # Abrir PDF automaticamente
        try:
            if sys.platform == "win32":
                os.startfile(str(arquivo))
        except Exception:
            pass

    def _reenviar_email(self):
        reg = self._selecionado()
        if not reg:
            return
        email = reg.get("email", "").strip()
        arquivo = reg.get("arquivo_pdf", "")
        if not validar_email(email):
            messagebox.showwarning("E-mail inválido",
                                    f"O e-mail '{email}' não é válido.\n"
                                    "Atualize a planilha e use 'Reimprimir'.",
                                    parent=self.win)
            return
        if not arquivo or not Path(arquivo).exists():
            messagebox.showwarning("Arquivo não encontrado",
                                    "PDF original não encontrado.\n"
                                    "Use 'Reimprimir' para gerar outro.",
                                    parent=self.win)
            return

        ok, msg = Envio.email(email, Path(arquivo), reg.get("nome", ""),
                                reg.get("matricula", ""), self.operador)
        if ok:
            self.historico.marcar_enviado(reg.get("codigo", ""), "email")
            messagebox.showinfo("Enviado", msg, parent=self.win)
            self._recarregar()
        else:
            messagebox.showerror("Falha", msg, parent=self.win)

    def _reenviar_whatsapp(self):
        reg = self._selecionado()
        if not reg:
            return
        tel = reg.get("telefone", "").strip()
        arquivo = reg.get("arquivo_pdf", "")
        if not normalizar_telefone(tel):
            messagebox.showwarning("Telefone inválido",
                                    f"O telefone '{tel}' não é válido.",
                                    parent=self.win)
            return
        if not arquivo or not Path(arquivo).exists():
            messagebox.showwarning("Arquivo não encontrado",
                                    "PDF original não encontrado.",
                                    parent=self.win)
            return

        ok, msg = Envio.whatsapp(tel, Path(arquivo), reg.get("nome", ""),
                                   reg.get("matricula", ""), self.operador)
        if ok:
            self.historico.marcar_enviado(reg.get("codigo", ""), "whatsapp")
            messagebox.showinfo("WhatsApp", msg, parent=self.win)
            self._recarregar()
        else:
            messagebox.showwarning("Atenção", msg, parent=self.win)

    def _abrir_pasta(self):
        try:
            if sys.platform == "win32":
                os.startfile(str(OUTPUT_DIR))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(OUTPUT_DIR)])
            else:
                subprocess.Popen(["xdg-open", str(OUTPUT_DIR)])
        except Exception as e:
            messagebox.showerror("Erro", str(e), parent=self.win)


# ============================================================
# JANELA PRINCIPAL
# ============================================================
class MainWindow:
    def __init__(self, seguranca: Seguranca):
        self.seg = seguranca
        self.operador = seguranca.operador()
        self.repo = RepositorioAssociados()
        self.gerador = GeradorCertificado()
        self.historico = HistoricoEmissoes(HISTORICO_CSV)
        self.arquivo_atual = None
        self.codigo_atual = None
        self.dados_atual = {}  # guarda nome/email/matr/tel do certificado atual

        self.root = tk.Tk()
        self.root.title("Emissor de Certificados · Mercado Central BH")
        self.root.configure(bg=IVORY)
        self.root.geometry("1000x750")
        self._centralizar(1000, 750)

        self._montar_ui()
        self._tentar_carregar_planilha_padrao()

        self.root.mainloop()

    def _centralizar(self, w, h):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ---------- UI ----------
    def _montar_ui(self):
        # ========== CABEÇALHO ==========
        header = tk.Frame(self.root, bg=NAVY, height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        # Logo
        try:
            from PIL import Image, ImageTk
            img = Image.open(LOGO_PATH).convert("RGBA")
            img.thumbnail((55, 55), Image.LANCZOS)
            self._logo_img = ImageTk.PhotoImage(img)
            tk.Label(header, image=self._logo_img, bg=NAVY).pack(side="left", padx=20, pady=12)
        except Exception:
            tk.Label(header, text="🏛", font=("Arial", 32),
                     bg=NAVY, fg=GOLD).pack(side="left", padx=20, pady=12)

        title_frame = tk.Frame(header, bg=NAVY)
        title_frame.pack(side="left", pady=12)
        tk.Label(title_frame, text="MERCADO CENTRAL",
                 font=("Georgia", 14, "bold"),
                 bg=NAVY, fg=WHITE).pack(anchor="w")
        tk.Label(title_frame, text="Emissor de Certificados de Regularidade",
                 font=("Calibri", 10),
                 bg=NAVY, fg=GOLD).pack(anchor="w")

        # Operador + botões à direita
        right = tk.Frame(header, bg=NAVY)
        right.pack(side="right", padx=20)
        tk.Label(right, text=f"Operador: {self.operador}",
                 font=("Calibri", 10, "italic"),
                 bg=NAVY, fg=SLATE_LT).pack(anchor="e", pady=(12, 2))
        btns_header = tk.Frame(right, bg=NAVY)
        btns_header.pack(anchor="e")
        tk.Button(btns_header, text="📂  Carregar planilha",
                  command=self._escolher_planilha,
                  bg=GOLD, fg=NAVY, font=("Calibri", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=4,
                  activebackground=GOLD_DEEP).pack(side="left", padx=(0, 6))
        tk.Button(btns_header, text="📋  Histórico",
                  command=self._abrir_historico,
                  bg=NAVY_DARK, fg=GOLD, font=("Calibri", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=4,
                  activebackground=NAVY, activeforeground=WHITE).pack(side="left", padx=(0, 6))
        tk.Button(btns_header, text="📊  Exportar base",
                  command=self._exportar_base,
                  bg=NAVY_DARK, fg=GOLD, font=("Calibri", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=4,
                  activebackground=NAVY, activeforeground=WHITE).pack(side="left", padx=(0, 6))
        tk.Button(btns_header, text="⚙  Configurações",
                  command=self._abrir_config,
                  bg=NAVY_DARK, fg=GOLD, font=("Calibri", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=4,
                  activebackground=NAVY, activeforeground=WHITE).pack(side="left")

        tk.Frame(self.root, bg=GOLD, height=3).pack(fill="x")

        # ========== CORPO ==========
        body = tk.Frame(self.root, bg=IVORY)
        body.pack(fill="both", expand=True, padx=30, pady=20)

        # ---------- SEÇÃO 1: BUSCA ----------
        sec1 = tk.LabelFrame(body, text="  1 · Buscar associado  ",
                             font=("Calibri", 11, "bold"),
                             bg=IVORY, fg=NAVY, bd=1, relief="solid",
                             labelanchor="nw", padx=20, pady=15)
        sec1.pack(fill="x", pady=(0, 15))

        tk.Label(sec1, text="Digite o nome ou matrícula (mínimo 2 caracteres):",
                 font=("Calibri", 10), bg=IVORY, fg=SLATE).pack(anchor="w", pady=(0, 5))

        self.autocomplete = AutocompleteEntry(
            sec1, self.repo, on_select=self._preencher_dados, bg=IVORY
        )
        self.autocomplete.pack(fill="x")

        self.lbl_total = tk.Label(sec1, text="",
                                   font=("Calibri", 9, "italic"),
                                   bg=IVORY, fg=SLATE_LT)
        self.lbl_total.pack(anchor="w", pady=(8, 0))

        # ---------- SEÇÃO 2: DADOS ----------
        sec2 = tk.LabelFrame(body, text="  2 · Dados do associado (editáveis)  ",
                             font=("Calibri", 11, "bold"),
                             bg=IVORY, fg=NAVY, bd=1, relief="solid",
                             labelanchor="nw", padx=20, pady=15)
        sec2.pack(fill="x", pady=(0, 15))

        grid = tk.Frame(sec2, bg=IVORY)
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        def _label(parent, text, r, c):
            tk.Label(parent, text=text, font=("Calibri", 10, "bold"),
                     bg=IVORY, fg=NAVY).grid(row=r, column=c, sticky="w",
                                              padx=(0, 10), pady=6)

        def _entry(parent, r, c, cs=1):
            var = tk.StringVar()
            e = tk.Entry(parent, textvariable=var, font=("Calibri", 11),
                         relief="solid", bd=1, bg=WHITE,
                         highlightthickness=1, highlightbackground=DIVIDER,
                         highlightcolor=GOLD)
            e.grid(row=r, column=c, columnspan=cs, sticky="ew", padx=(0, 15),
                   pady=6, ipady=5)
            return var, e

        _label(grid, "Nome completo:", 0, 0)
        self.v_nome, _ = _entry(grid, 0, 1, cs=3)

        _label(grid, "Matrícula:", 1, 0)
        self.v_matr, _ = _entry(grid, 1, 1)
        _label(grid, "Telefone:", 1, 2)
        self.v_tel, _ = _entry(grid, 1, 3)

        _label(grid, "E-mail:", 2, 0)
        self.v_email, _ = _entry(grid, 2, 1, cs=3)

        _label(grid, "Data Assembleia:", 3, 0)
        self.v_assem = tk.StringVar(value=self.seg.data_assembleia_padrao())
        e_assem = tk.Entry(grid, textvariable=self.v_assem, font=("Calibri", 11),
                           relief="solid", bd=1, bg=WHITE,
                           highlightthickness=1, highlightbackground=DIVIDER,
                           highlightcolor=GOLD)
        e_assem.grid(row=3, column=1, columnspan=3, sticky="ew",
                     padx=(0, 15), pady=6, ipady=5)

        # ---------- SEÇÃO 3: AÇÕES ----------
        sec3 = tk.LabelFrame(body, text="  3 · Gerar e enviar  ",
                             font=("Calibri", 11, "bold"),
                             bg=IVORY, fg=NAVY, bd=1, relief="solid",
                             labelanchor="nw", padx=20, pady=15)
        sec3.pack(fill="x", pady=(0, 15))

        botoes = tk.Frame(sec3, bg=IVORY)
        botoes.pack(fill="x")

        self.btn_gerar = tk.Button(
            botoes, text="📄  Gerar Certificado",
            command=self._gerar_certificado,
            bg=NAVY, fg=WHITE, font=("Calibri", 11, "bold"),
            relief="flat", cursor="hand2", padx=20, pady=10,
            activebackground=NAVY_DARK, activeforeground=WHITE
        )
        self.btn_gerar.pack(side="left", padx=(0, 10))

        self.btn_email = tk.Button(
            botoes, text="📧  Enviar por E-mail",
            command=self._enviar_email, state="disabled",
            bg=SUCCESS, fg=WHITE, font=("Calibri", 11, "bold"),
            relief="flat", cursor="hand2", padx=20, pady=10,
            disabledforeground=SLATE_LT
        )
        self.btn_email.pack(side="left", padx=(0, 10))

        self.btn_zap = tk.Button(
            botoes, text="📱  Enviar por WhatsApp",
            command=self._enviar_whatsapp, state="disabled",
            bg="#25D366", fg=WHITE, font=("Calibri", 11, "bold"),
            relief="flat", cursor="hand2", padx=20, pady=10,
            disabledforeground=SLATE_LT
        )
        self.btn_zap.pack(side="left", padx=(0, 10))

        self.btn_abrir = tk.Button(
            botoes, text="👁  Abrir PDF",
            command=self._abrir_pdf, state="disabled",
            bg=IVORY, fg=NAVY, font=("Calibri", 10),
            relief="solid", bd=1, cursor="hand2", padx=15, pady=10,
            disabledforeground=SLATE_LT
        )
        self.btn_abrir.pack(side="left")

        self.btn_limpar = tk.Button(
            botoes, text="Limpar",
            command=self._limpar,
            bg=IVORY, fg=SLATE, font=("Calibri", 10),
            relief="flat", cursor="hand2", padx=10, pady=10
        )
        self.btn_limpar.pack(side="right")

        # ---------- STATUS ----------
        self.status_frame = tk.Frame(body, bg=IVORY, height=40)
        self.status_frame.pack(fill="x")

        self.lbl_status = tk.Label(
            self.status_frame, text="Pronto. Carregue a planilha de associados para começar.",
            font=("Calibri", 10, "italic"),
            bg=IVORY, fg=SLATE, anchor="w"
        )
        self.lbl_status.pack(fill="x", pady=5)

        # ---------- RODAPÉ ----------
        footer = tk.Frame(self.root, bg=NAVY, height=40)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        tk.Label(footer,
                 text="Desenvolvido por Artur Gabriel Oliveira da Silva · arturg_oliveira@outlook.com",
                 font=("Calibri", 9, "italic"),
                 bg=NAVY, fg=GOLD).pack(pady=10)

    # ---------- AÇÕES ----------
    def _tentar_carregar_planilha_padrao(self):
        if PLANILHA_PATH.exists():
            ok, msg = self.repo.carregar(PLANILHA_PATH)
            if ok:
                self.lbl_total.config(
                    text=f"✓ Planilha padrão carregada: {self.repo.total()} associados.")
                self.lbl_status.config(
                    text="Planilha carregada. Digite o nome ou matrícula para buscar.")
            else:
                self.lbl_status.config(text=f"Erro ao carregar planilha padrão: {msg}",
                                       fg=DANGER)

    def _escolher_planilha(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha de associados",
            filetypes=[("Planilhas Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
            initialdir=str(BASE_DIR)
        )
        if not path:
            return
        ok, msg = self.repo.carregar(Path(path))
        if ok:
            self.lbl_total.config(
                text=f"✓ {self.repo.total()} associados carregados de {Path(path).name}.")
            self.lbl_status.config(
                text="Planilha carregada. Digite o nome ou matrícula para buscar.",
                fg=SLATE)
        else:
            messagebox.showerror("Erro ao carregar", msg)

    def _abrir_config(self):
        """Abre a janela de configurações do certificado."""
        ConfigWindow(self.root, self.seg, on_close=self._config_salva)

    def _abrir_historico(self):
        """Abre a janela de histórico de emissões."""
        HistoricoWindow(self.root, self.seg, self.historico,
                         self.gerador, self.operador)

    def _exportar_base(self):
        """
        Exporta a base de associados enriquecida com dados das emissões.

        Comportamento:
          - Pega todos os associados da planilha carregada
          - Para cada um, busca no histórico o e-mail/telefone mais recente usado
          - Gera planilha com a base ATUALIZADA (planilha ∪ histórico)
        """
        if not self.repo.carregado or self.repo.total() == 0:
            messagebox.showwarning(
                "Sem dados",
                "Nenhuma planilha está carregada.\n"
                "Clique em 'Carregar planilha' primeiro."
            )
            return

        sugestao = f"base_associados_atualizacao_{datetime.now():%Y%m%d}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Salvar planilha de atualização",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
            initialfile=sugestao,
            initialdir=str(BASE_DIR)
        )
        if not path:
            return

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

            # ==========================================================
            # MERGE: planilha base + últimos dados usados em emissões
            # ==========================================================
            # Indexar histórico por matrícula, pegando a emissão MAIS RECENTE
            emissoes_por_matr = {}  # matricula -> dados mais recentes
            for emissao in self.historico.listar():  # já vem ordenado desc
                matr = str(emissao.get("matricula", "")).strip()
                if matr and matr not in emissoes_por_matr:
                    emissoes_por_matr[matr] = emissao

            # Aplicar merge: email/telefone do histórico sobrescreve o da planilha
            # SE o do histórico estiver preenchido (operador digitou manualmente)
            registros_enriquecidos = []
            atualizados = 0
            for reg in self.repo.registros:
                matr = str(reg.get("matricula", "")).strip()
                email_final = reg.get("email", "").strip()
                tel_final = reg.get("telefone", "").strip()
                atualizado_por_emissao = False

                if matr in emissoes_por_matr:
                    hist = emissoes_por_matr[matr]
                    email_hist = hist.get("email", "").strip()
                    tel_hist = hist.get("telefone", "").strip()

                    # Substitui se histórico tem dado novo OU diferente
                    if email_hist and email_hist != email_final:
                        email_final = email_hist
                        atualizado_por_emissao = True
                    if tel_hist and tel_hist != tel_final:
                        tel_final = tel_hist
                        atualizado_por_emissao = True

                if atualizado_por_emissao:
                    atualizados += 1

                registros_enriquecidos.append({
                    "matricula": reg.get("matricula", ""),
                    "nome":      reg.get("nome", ""),
                    "email":     email_final,
                    "telefone":  tel_final,
                    "atualizado": atualizado_por_emissao,
                })

            # ==========================================================
            # Gerar planilha Excel formatada
            # ==========================================================
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Associados"

            # Título
            ws.merge_cells("A1:D1")
            ws["A1"] = "MERCADO CENTRAL DE BELO HORIZONTE"
            ws["A1"].font = Font(name="Georgia", size=14, bold=True, color="FFFFFF")
            ws["A1"].fill = PatternFill("solid", fgColor="1A2332")
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 28

            ws.merge_cells("A2:D2")
            ws["A2"] = f"Base de Associados — atualização em {datetime.now():%d/%m/%Y %H:%M}"
            ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="9C7F3E")
            ws["A2"].alignment = Alignment(horizontal="center")
            ws.row_dimensions[2].height = 20

            # Cabeçalho
            headers = ["Matrícula", "Nome", "E-mail", "Telefone/WhatsApp"]
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=4, column=col, value=h)
                c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1A2332")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = Border(
                    top=Side(style="thin", color="C9A961"),
                    bottom=Side(style="medium", color="C9A961"),
                )
            ws.row_dimensions[4].height = 26

            # Dados (com destaque verde nas linhas atualizadas via histórico)
            thin = Side(style="thin", color="E2DDD3")
            border_data = Border(left=thin, right=thin, top=thin, bottom=thin)

            for idx, reg in enumerate(registros_enriquecidos, start=5):
                ws.cell(row=idx, column=1, value=reg["matricula"])
                ws.cell(row=idx, column=2, value=reg["nome"])
                ws.cell(row=idx, column=3, value=reg["email"])
                ws.cell(row=idx, column=4, value=reg["telefone"])
                for col in range(1, 5):
                    cell = ws.cell(row=idx, column=col)
                    cell.font = Font(name="Calibri", size=10)
                    cell.border = border_data
                    cell.alignment = Alignment(vertical="center",
                                                horizontal="left" if col in (2, 3) else "center")

                # Destacar linhas atualizadas pelo histórico (verde claro)
                if reg.get("atualizado"):
                    for col in range(1, 5):
                        ws.cell(row=idx, column=col).fill = PatternFill(
                            "solid", fgColor="E8F5E9")  # verde claro
                elif idx % 2 == 1:
                    for col in range(1, 5):
                        ws.cell(row=idx, column=col).fill = PatternFill(
                            "solid", fgColor="F8F5F0")

            # Larguras
            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 42
            ws.column_dimensions["C"].width = 34
            ws.column_dimensions["D"].width = 22

            ws.freeze_panes = "A5"

            # Aba Instruções
            ws_info = wb.create_sheet("Instruções")
            ws_info["A1"] = "Instruções para atualização da base"
            ws_info["A1"].font = Font(name="Georgia", size=14, bold=True, color="1A2332")

            instrucoes = [
                "",
                "Esta planilha contém a base de associados atualizada com os dados mais",
                "recentes coletados durante as emissões de certificado.",
                "",
                "COMO USAR:",
                "",
                "1. Revise os dados de cada associado na aba 'Associados'",
                "2. Linhas com FUNDO VERDE foram atualizadas automaticamente —",
                "   o operador digitou um e-mail ou telefone novo durante a emissão",
                "3. Corrija, adicione ou remova dados conforme necessário",
                "4. Para ADICIONAR associado novo: insira linha abaixo da última",
                "5. Para REMOVER associado: apague a linha inteira",
                "6. NÃO altere os nomes das colunas",
                "7. Salve como .xlsx e substitua o arquivo 'planilha.xlsx' na pasta do programa",
                "",
                "IMPORTANTE: mantenha o formato do arquivo (.xlsx).",
                "",
                f"Total de associados: {len(registros_enriquecidos)}",
                f"Atualizados por emissões recentes: {atualizados}",
                f"Gerado em: {datetime.now():%d/%m/%Y %H:%M}",
                f"Operador: {self.operador}",
            ]
            for i, txt in enumerate(instrucoes, start=3):
                c = ws_info.cell(row=i, column=1, value=txt)
                c.font = Font(name="Calibri", size=11)
                if txt.startswith("IMPORTANTE") or txt.startswith("COMO USAR"):
                    c.font = Font(name="Calibri", size=11, bold=True,
                                   color="C53030" if "IMPORTANTE" in txt else "1A2332")
                elif "VERDE" in txt:
                    c.fill = PatternFill("solid", fgColor="E8F5E9")

            ws_info.column_dimensions["A"].width = 85

            wb.save(path)

            # Mensagem com resumo do merge
            msg_atualizados = ""
            if atualizados > 0:
                msg_atualizados = (f"\n\n✓ {atualizados} registro(s) foram atualizados "
                                    f"com dados novos capturados nas emissões.\n"
                                    f"(Linhas com fundo VERDE na planilha)")

            if messagebox.askyesno(
                "Planilha exportada",
                f"Planilha gerada com {len(registros_enriquecidos)} associados.{msg_atualizados}\n\n"
                f"Deseja abrir agora para revisão?"
            ):
                try:
                    if sys.platform == "win32":
                        os.startfile(path)
                    elif sys.platform == "darwin":
                        subprocess.Popen(["open", path])
                    else:
                        subprocess.Popen(["xdg-open", path])
                except Exception:
                    pass

            registrar_auditoria(self.operador, "", "", path, "exportacao",
                                 "SUCESSO", f"{len(registros_enriquecidos)} regs, {atualizados} atualizados")
            log.info(f"Base exportada: {path} ({atualizados} atualizados do histórico)")

        except Exception as e:
            log.exception("Erro ao exportar base")
            messagebox.showerror("Erro",
                                   f"Falha ao exportar:\n{e}")

    def _config_salva(self):
        """Callback após salvar configurações — atualiza UI."""
        self.v_assem.set(self.seg.data_assembleia_padrao())
        self.lbl_status.config(
            text="✓ Configurações atualizadas. Novos certificados usarão os dados atualizados.",
            fg=SUCCESS
        )

    def _preencher_dados(self, registro):
        self.v_nome.set(registro["nome"])
        self.v_matr.set(registro["matricula"])
        self.v_email.set(registro["email"])
        self.v_tel.set(registro["telefone"])
        self.lbl_status.config(
            text=f"Associado selecionado: {registro['nome']}. "
                 f"Revise os dados e clique em Gerar Certificado.",
            fg=NAVY
        )

    def _gerar_certificado(self):
        nome = self.v_nome.get().strip()
        matr = self.v_matr.get().strip()
        assem = self.v_assem.get().strip()

        if not nome or len(nome) < 3:
            messagebox.showwarning("Atenção", "Informe o nome do associado.")
            return
        if not matr:
            messagebox.showwarning("Atenção", "Informe a matrícula.")
            return
        if not assem:
            messagebox.showwarning("Atenção", "Informe a data da Assembleia.")
            return

        self.lbl_status.config(text="Gerando certificado...", fg=NAVY)
        self.root.update()

        arquivo, codigo, erro = self.gerador.gerar(
            nome, matr, assem,
            signatario_nome=self.seg.signatario_nome(),
            signatario_cargo=self.seg.signatario_cargo(),
            assinatura_path=self.seg.signatario_assinatura_path(),
        )
        if not arquivo:
            messagebox.showerror("Erro ao gerar", erro or "Falha desconhecida.")
            self.lbl_status.config(text="Falha na geração.", fg=DANGER)
            return

        self.arquivo_atual = arquivo
        self.codigo_atual = codigo
        self.dados_atual = {
            "codigo": codigo,
            "matricula": matr,
            "nome": nome,
            "email": self.v_email.get().strip(),
            "telefone": self.v_tel.get().strip(),
            "data_assembleia": assem,
            "arquivo_pdf": str(arquivo),
            "operador": self.operador,
            "enviado_email": "",
            "enviado_whatsapp": "",
        }
        self.historico.registrar(self.dados_atual)
        self.lbl_status.config(
            text=f"✓ Certificado gerado: {arquivo.name}  │  Código: {codigo}",
            fg=SUCCESS
        )
        self.btn_email.config(state="normal")
        self.btn_zap.config(state="normal")
        self.btn_abrir.config(state="normal")
        registrar_auditoria(self.operador, matr, nome, "", "geracao",
                            "SUCESSO", codigo)

    def _enviar_email(self):
        email = self.v_email.get().strip()
        if not validar_email(email):
            messagebox.showwarning("E-mail inválido",
                                   "O e-mail informado não é válido.")
            return

        if not SMTP_USER or not SMTP_PASS:
            messagebox.showerror(
                "SMTP não configurado",
                "As credenciais de e-mail não estão configuradas.\n\n"
                "Edite o arquivo .env na pasta do programa com:\n"
                "SMTP_USER=seu_email@...\nSMTP_PASS=sua_senha_de_app"
            )
            return

        self.lbl_status.config(text="Enviando e-mail...", fg=NAVY)
        self.root.update()
        ok, msg = Envio.email(email, self.arquivo_atual, self.v_nome.get().strip(),
                              self.v_matr.get().strip(), self.operador)
        if ok:
            if self.codigo_atual:
                self.historico.marcar_enviado(self.codigo_atual, "email")
            messagebox.showinfo("Enviado", msg)
            self.lbl_status.config(text=f"✓ {msg}", fg=SUCCESS)
        else:
            messagebox.showerror("Falha no envio", msg)
            self.lbl_status.config(text=f"Falha: {msg}", fg=DANGER)

    def _enviar_whatsapp(self):
        tel = self.v_tel.get().strip()
        if not normalizar_telefone(tel):
            messagebox.showwarning("Telefone inválido",
                                    "Informe um telefone com DDD válido.")
            return
        ok, msg = Envio.whatsapp(tel, self.arquivo_atual,
                                  self.v_nome.get().strip(),
                                  self.v_matr.get().strip(),
                                  self.operador)
        if ok:
            if self.codigo_atual:
                self.historico.marcar_enviado(self.codigo_atual, "whatsapp")
            messagebox.showinfo("WhatsApp", msg)
            self.lbl_status.config(text=f"✓ {msg}", fg=SUCCESS)
        else:
            messagebox.showwarning("Atenção", msg)

    def _abrir_pdf(self):
        if not self.arquivo_atual:
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(self.arquivo_atual))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(self.arquivo_atual)])
            else:
                subprocess.Popen(["xdg-open", str(self.arquivo_atual)])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o PDF:\n{e}")

    def _limpar(self):
        self.autocomplete.set("")
        self.v_nome.set("")
        self.v_matr.set("")
        self.v_email.set("")
        self.v_tel.set("")
        self.v_assem.set(self.seg.data_assembleia_padrao())
        self.arquivo_atual = None
        self.codigo_atual = None
        self.btn_email.config(state="disabled")
        self.btn_zap.config(state="disabled")
        self.btn_abrir.config(state="disabled")
        self.lbl_status.config(
            text="Pronto para o próximo associado.",
            fg=SLATE
        )
        self.autocomplete.focus()


# ============================================================
# INICIALIZAÇÃO
# ============================================================
def verificar_arquivos() -> list[str]:
    problemas = []
    if not MODELO_PATH.exists():
        problemas.append(f"Modelo do certificado não encontrado: {MODELO_PATH.name}")
    if not LOGO_PATH.exists():
        problemas.append(f"Logo não encontrada: {LOGO_PATH.name}")
    return problemas


def main():
    problemas = verificar_arquivos()
    if problemas:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror(
            "Arquivos faltando",
            "O programa não pode iniciar:\n\n• " + "\n• ".join(problemas) +
            "\n\nVerifique a pasta do programa e tente novamente."
        )
        return

    seg = Seguranca()
    LoginWindow(seg, on_success=lambda _operador: MainWindow(seg))


if __name__ == "__main__":
    main()
