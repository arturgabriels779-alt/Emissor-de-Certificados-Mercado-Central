<div align="center">

<img src="logo.png" width="120" alt="Mercado Central BH Logo" />

# Emissor de Certificados de Regularidade Associativa

### Mercado Central de Belo Horizonte

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?style=flat-square&logo=windows&logoColor=white)](https://microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)
[![Version](https://img.shields.io/badge/Version-5.0-C9A961?style=flat-square)](CHANGELOG.md)

*Sistema desktop para emissão automatizada de certificados associativos com envio por e-mail e WhatsApp.*

</div>

---

## Visão Geral

Aplicativo desenvolvido para a administração do **Mercado Central de Belo Horizonte** (fundado em 1929), que digitaliza o processo de emissão dos Certificados de Regularidade Associativa. Substitui um fluxo 100% manual — busca em planilha, preenchimento em Word, impressão, entrega — por um processo de **15 segundos por emissão**, com PDF gerado automaticamente, QR Code de autenticidade e envio direto ao associado.

> **Contexto real:** O Mercado Central de BH tem ~600 associados. Cada um precisa de um Certificado de Regularidade para participar das Assembleias Gerais Ordinárias. Este sistema centraliza, automatiza e rastreia todo esse processo.

---

## Funcionalidades

| Módulo | Recursos |
|--------|----------|
| **Emissão** | Autocomplete inteligente por nome ou matrícula · Preenchimento automático de dados · PDF gerado com QR Code único · Data e assinatura do signatário configuráveis |
| **Envio** | E-mail com PDF anexado (SMTP) · WhatsApp via link `wa.me` com mensagem pronta · Registro de canal de envio por emissão |
| **Histórico** | Tabela pesquisável (texto + intervalo de datas) · Indicação de envio por e-mail e WhatsApp · Reimprimir com novo código · Reenviar por qualquer canal |
| **Base de dados** | Carrega planilha `.xlsx` com detecção automática de colunas · Aceita formato mínimo (só nome + matrícula) · Exporta base enriquecida com dados coletados nas emissões |
| **Segurança** | PIN de acesso com hash PBKDF2-SHA256 + salt · Recuperação de PIN por código de 6 dígitos via e-mail · Bloqueio após 3 tentativas · Log de auditoria CSV |
| **Configurações** | Nome e cargo do signatário · Imagem de assinatura digital (PNG com transparência) · Data padrão da Assembleia · Abrir template Word diretamente |

---

## Screenshots

> *(Adicione aqui prints das telas principais após o deploy em produção)*

---

## Stack Técnica

```
Python 3.10+
├── tkinter          — Interface gráfica nativa (sem dependência de framework UI externo)
├── docxtpl          — Renderização do template Word com Jinja2 (preserva formatação)
├── docx2pdf         — Conversão DOCX → PDF via MS Word COM (Windows)
├── qrcode           — Geração do QR Code de autenticidade
├── pandas           — Leitura e manipulação da planilha de associados
├── openpyxl         — Exportação da planilha enriquecida
├── python-dotenv    — Isolamento de credenciais SMTP
├── smtplib          — Envio de e-mail com anexo (stdlib)
└── Pillow           — Processamento da logo com transparência
```

**Arquitetura:** aplicativo monolítico single-file (`app.py`, ~2900 linhas) com 12 classes organizadas em camadas:

```
Seguranca · RepositorioAssociados · GeradorCertificado · Envio · HistoricoEmissoes
                                    ↓
AutocompleteEntry · LoginWindow · RecuperacaoWindow · TrocarPinWindow
ConfigWindow · HistoricoWindow · MainWindow
```

---

## Instalação

### Pré-requisitos

- Windows 10 ou 11
- Microsoft Word instalado *(para conversão PDF via COM; fallback: LibreOffice)*
- Conexão com a internet *(apenas na primeira execução, para instalar dependências)*

### Setup Automático

```bash
# 1. Clone o repositório
git clone https://github.com/seu-usuario/emissor-certificados-mercadocentral.git
cd emissor-certificados-mercadocentral

# 2. Dê duplo-clique em IniciarPrograma.bat
#    — Se Python não estiver instalado, o .bat oferece instalação automática
#    — Instala todas as dependências no primeiro uso (~3 min)
#    — Abre o app
```

### Setup Manual (alternativa)

```bash
pip install -r requirements.txt
python app.py
```

---

## Configuração

### 1. Arquivo `.env` (credenciais de e-mail)

Renomeie `env_exemplo.txt` para `.env` e preencha:

```env
# Microsoft 365 / Outlook
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_USER=secretaria@mercadocentral.com.br
SMTP_PASS=sua_senha_aqui
SMTP_FROM=Mercado Central BH - Secretaria

# Gmail (requer Senha de App — não a senha normal)
# SMTP_HOST=smtp.gmail.com
# SMTP_PORT=587
# SMTP_USER=seu@gmail.com
# SMTP_PASS=xxxx xxxx xxxx xxxx   ← Senha de App do Google
```

### 2. Planilha de associados

O sistema aceita dois formatos:

**Formato completo:**
| matricula | nome | email | telefone |
|-----------|------|-------|----------|
| 0001 | Antônio Silva | antonio@email.com | (31) 99111-1111 |

**Formato mínimo** *(e-mail e telefone preenchidos pelo operador na emissão):*
| ASSOCIADO | MATRICULA |
|-----------|-----------|
| ACÁCIO ALVES DUARTE | 583 |

O sistema detecta automaticamente variações nos nomes das colunas: `ASSOCIADO`, `NOME`, `MATRICULA`, `MATRÍCULA`, `E-MAIL`, `EMAIL`, `CELULAR`, `WHATSAPP`, etc.

### 3. Template do certificado

Edite `modelo_certificado.docx` pelo Word para personalizar o layout. Preserve os placeholders Jinja2:

| Placeholder | Conteúdo |
|---|---|
| `{{ nome }}` | Nome do associado em MAIÚSCULAS |
| `{{ matricula }}` | Matrícula |
| `{{ data_emissao }}` | Data por extenso (gerada automaticamente) |
| `{{ data_assembleia }}` | Data da AGO (configurável) |
| `{{ signatario_nome }}` | Nome do signatário |
| `{{ signatario_cargo }}` | Cargo do signatário |
| `{{ assinatura_img }}` | Imagem de assinatura OU linha em branco |
| `{{ qr_code }}` | QR Code de autenticidade |
| `{{ codigo_verificacao }}` | Código único alfanumérico |

---

## Estrutura do Repositório

```
emissor-certificados-mercadocentral/
│
├── app.py                        # Aplicação principal (~2900 linhas, 12 classes)
├── modelo_certificado.docx       # Template Word com placeholders Jinja2
├── logo.png                      # Logo do Mercado (PNG com transparência)
├── planilha.xlsx                 # Planilha de exemplo com 5 associados fictícios
│
├── IniciarPrograma.bat           # Launcher Windows com auto-install Python/deps
├── requirements.txt              # Dependências Python
├── env_exemplo.txt               # Template do arquivo .env
│
├── certificados_gerados/         # PDFs gerados (criado automaticamente, no .gitignore)
└── logs/                         # Logs técnicos e auditoria (criado automaticamente)
    ├── app_YYYYMMDD.log
    ├── auditoria_envios.csv
    └── historico_emissoes.csv
```

---

## Segurança

- **PIN de acesso** armazenado com PBKDF2-HMAC-SHA256 + salt aleatório (100.000 iterações)
- **Credenciais SMTP** em arquivo `.env` isolado do código-fonte — nunca commitar
- **Bloqueio automático** após 3 tentativas de PIN erradas (5 minutos)
- **Recuperação de PIN** por código de 6 dígitos com validade de 15 minutos, enviado ao e-mail cadastrado
- **Auditoria CSV** com timestamp, operador, matrícula, canal e status de cada envio
- **Sanitização** de inputs antes de renderização no template (previne injeção Jinja2)
- **Validação** de e-mail (regex RFC 5322) e telefone antes do envio

> **⚠️ Atenção:** Adicione ao `.gitignore` os arquivos `.env`, `config.json`, `planilha.xlsx` (dados reais) e a pasta `certificados_gerados/`.

---

## `.gitignore` recomendado

```gitignore
# Credenciais e dados sensíveis
.env
config.json

# Planilha com dados reais de associados
planilha.xlsx

# Arquivos gerados em tempo de execução
certificados_gerados/
logs/
assinatura.png
assinatura.jpg

# Python
__pycache__/
*.pyc
*.pyo
*.pyd
.Python
*.egg-info/
dist/
build/
.venv/
venv/
env/

# Windows
Thumbs.db
desktop.ini
```

---

## Fluxo do WhatsApp

O WhatsApp Business API (que permite envio de arquivos automaticamente) exige aprovação da Meta e tem custo por mensagem. A solução adotada usa `wa.me` com link de conversa pré-preenchida:

1. O sistema abre o WhatsApp Web com uma mensagem personalizada pronta
2. Simultaneamente, abre a pasta do certificado no Explorador de Arquivos
3. O operador arrasta o PDF para o chat e envia com 1 clique

Para quem quiser evoluir para envio com anexo automático, a arquitetura suporta integração com a [API Cloud da Meta](https://developers.facebook.com/docs/whatsapp/cloud-api) como evolução futura.

---

## Roadmap

- [ ] **v1.1** — Integração com Alterdata ERP (puxar status de regularidade direto do banco de dados)
- [ ] **v1.2** — Modo lote: emitir e enviar para múltiplos associados de uma vez
- [ ] **v1.3** — Dashboard de emissões com gráficos (integração com Power BI ou panel interno)
- [ ] **v2.0** — Assinatura digital ICP-Brasil com validade jurídica plena
- [ ] **v2.1** — WhatsApp Business API (Meta Cloud) para envio de anexo sem interação manual
- [ ] **v3.0** — Versão web (FastAPI + React) — alinhamento com stack já em desenvolvimento no Mercado Central

---

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para detalhes.

---

## Autor

<table>
  <tr>
    <td align="center">
      <strong>Artur Gabriel Oliveira da Silva</strong><br>
      Administração · Desenvolvimento de Sistemas<br>
      Mercado Central de Belo Horizonte<br>
      <a href="mailto:arturg_oliveira@outlook.com">arturg_oliveira@outlook.com</a><br>
      <a href="https://github.com/seu-usuario">GitHub</a>
    </td>
  </tr>
</table>

---

<div align="center">

*Desenvolvido com 🏛 para o Mercado Central de Belo Horizonte — desde 1929*

</div># Emissor-de-Certificados-Mercado-Central
