# Emissor de Certificados — Mercado Central BH

**Versão 4.0**
**Desenvolvido por:** Artur Gabriel Oliveira da Silva
**Contato:** arturg_oliveira@outlook.com

---

## Novidades da v4.0

- 🖼 **Logo corrigida** — fundo transparente de verdade, sem efeito xadrez cinza no certificado
- 📗 **Planilha flexível** — agora aceita planilha com apenas `nome` + `matrícula` (os dados de e-mail/WhatsApp podem ser digitados pelo operador no momento da emissão)
- 🔀 **Exportar base "inteligente"** — combina os dados da planilha carregada com os e-mails e telefones que o operador digitou durante as emissões. Linhas atualizadas ficam destacadas em verde na planilha gerada
- 📋 **Histórico com colunas claras** — colunas "E-mail enviado" e "WhatsApp enviado" agora mostram "✓ Sim" ou "—", não ícones minúsculos
- ⚙ **Botão "Abrir modelo no Word"** — agora fixo no rodapé da tela de Configurações, sempre acessível
- 🔍 **Reconhecimento automático de colunas** — "ASSOCIADO"/"NOME", "MATRICULA"/"ID", "E-MAIL"/"EMAIL", "CELULAR"/"WHATSAPP"/"TELEFONE"

---

## O que este programa faz

Emite **Certificados de Regularidade Associativa** em PDF, com:
- QR Code de autenticidade único
- Envio por e-mail (com anexo) ou WhatsApp (com aviso)
- Autocomplete inteligente na busca de associados
- Editor visual de signatário, cargo e assinatura digital
- Acesso protegido por PIN com recuperação por e-mail
- Histórico pesquisável de todos os certificados já emitidos
- Exportação da base atualizada para Excel

---

## Primeira instalação

### 1. Extrair o ZIP em um local fixo

```
C:\MercadoCentral\EmissorCertificados\
```

### 2. Duplo-clique em `IniciarPrograma.bat`

- Verifica se o Python está instalado. Se NÃO estiver, pergunta se quer instalar automaticamente (responda **S**)
- Instala as bibliotecas necessárias (só na primeira vez, ~3 minutos)
- Abre o app

### 3. Primeiro acesso

- Nome do operador
- PIN de 4 a 8 dígitos
- **E-mail para recuperação** (muito recomendado preencher — se esquecer o PIN, recupera por aqui)

### 4. Configurar o e-mail que envia os certificados

1. Abra `env_exemplo.txt`, renomeie para `.env`
2. Edite com as credenciais do e-mail institucional do Mercado
3. Se usar Gmail: criar Senha de App em https://myaccount.google.com/apppasswords

### 5. Planilha de associados

O arquivo `planilha.xlsx` pode vir em **dois formatos**:

**Formato completo** (todos os dados):
| matricula | nome | email | telefone |
|-----------|------|-------|----------|
| 0001 | Antônio Silva | antonio@email.com | (31) 99111-1111 |

**Formato mínimo** (só o que o Mercado tem hoje):
| ASSOCIADO | MATRICULA |
|-----------|-----------|
| ACÁCIO ALVES DUARTE | 583 |
| ADA ORLANDI MALVICINO VIEIRA | 1705 |

No formato mínimo, os campos de e-mail e telefone aparecem vazios no app e **o operador preenche manualmente** no momento da emissão. Esses dados são **memorizados no histórico** para uso futuro.

O programa aceita nomes de coluna variados: `ASSOCIADO`, `NOME`, `NOME_COMPLETO`, `MATRICULA`, `MATRÍCULA`, `MATRIC`, `ID`, `E-MAIL`, `EMAIL`, `TELEFONE`, `CELULAR`, `WHATSAPP`, etc.

### 6. Personalizar o certificado

Na tela principal, clique em **⚙ Configurações**:
- Nome e cargo do signatário
- Imagem de assinatura digital (PNG com fundo transparente)
- Data padrão da Assembleia
- E-mail de recuperação de PIN

Para personalizar o documento visualmente, clique em **📝 Abrir modelo no Word** no rodapé da tela de Configurações.

---

## Uso diário

### Emitir um certificado

1. Duplo-clique em `IniciarPrograma.bat`
2. Digite o PIN
3. Começe a digitar o nome ou matrícula — sugestões aparecem
4. Selecione o associado
5. Revise/complete os dados (se o e-mail ou telefone vierem vazios, digite aqui)
6. Clique em **Gerar Certificado**
7. Envie por **E-mail** ou **WhatsApp**

### Consultar emissões anteriores

**📋 Histórico** no header. Exibe tabela com:
- Data/Hora · Código · Matrícula · Nome · E-mail
- **"E-mail enviado"** → "✓ Sim" ou "—"
- **"WhatsApp enviado"** → "✓ Sim" ou "—"
- Operador

Filtros por texto (nome/matrícula/código) e data (AAAA-MM-DD).

Ações por item: **Abrir PDF**, **Reimprimir** (novo código), **Reenviar E-mail**, **Reenviar WhatsApp**.

### Exportar base enriquecida

**📊 Exportar base** no header. Gera Excel com **todos os associados da planilha + dados de e-mail/telefone que foram digitados no app**. Linhas com fundo **verde** = foram atualizadas com dados novos. É o arquivo ideal para:
- Entregar à equipe para revisão
- Substituir a planilha antiga (salvar sobre `planilha.xlsx`)
- Arquivar backup periódico

### Segurança

- **Trocar PIN**: Configurações → Segurança → **🔑 Trocar PIN agora**
- **Esqueci meu PIN**: link na tela de login (só aparece se você cadastrou e-mail de recuperação). Recebe código de 6 dígitos por e-mail, valida e cria novo PIN.

---

## Fluxo do WhatsApp

WhatsApp não permite anexar arquivos automaticamente sem API paga. Portanto:
- O programa abre o chat da pessoa no WhatsApp Web com uma mensagem pronta
- E abre a pasta do certificado numa janela do Windows
- Você arrasta o PDF para o chat e envia

---

## Estrutura de arquivos

```
EmissorCertificados/
├── app.py                         ← Programa principal
├── IniciarPrograma.bat            ← Duplo-clique para abrir
├── modelo_certificado.docx        ← Template editável
├── logo.png                       ← Logo com transparência
├── planilha.xlsx                  ← Base de associados
├── .env                           ← Credenciais de e-mail
├── requirements.txt               ← Dependências Python
├── config.json                    ← PIN + configurações (gerado pelo app)
├── assinatura.png                 ← Assinatura digital (opcional)
├── certificados_gerados/          ← PDFs ficam aqui
└── logs/
    ├── app_YYYYMMDD.log               ← Log técnico diário
    ├── auditoria_envios.csv           ← Todos os envios
    └── historico_emissoes.csv         ← Histórico do app
```

---

## Personalização avançada do certificado

Em **⚙ Configurações → 📝 Abrir modelo no Word**. Preserve os placeholders:

| Placeholder | Significado |
|---|---|
| `{{ nome }}` | Nome do associado (MAIÚSCULAS) |
| `{{ matricula }}` | Matrícula |
| `{{ data_emissao }}` | Data de hoje (por extenso) |
| `{{ data_assembleia }}` | Data da AGO |
| `{{ signatario_nome }}` | Nome do signatário (configurável) |
| `{{ signatario_cargo }}` | Cargo (configurável) |
| `{{ assinatura_img }}` | Imagem da rubrica OU linha em branco |
| `{{ codigo_verificacao }}` | Código único do certificado |
| `{{ qr_code }}` | QR Code |

---

## Problemas comuns

| Problema | Solução |
|----------|---------|
| Logo aparece com fundo cinza/xadrez | Verifique se `logo.png` é a versão v4 (tem transparência real). Se persistir, regenere o modelo |
| "Planilha não tem as colunas necessárias" | Planilha precisa ter, no mínimo, colunas de matrícula e nome. Veja a seção "Planilha de associados" |
| Histórico não atualiza após emissão | Feche e reabra a janela do histórico, ou o app |
| Exportar base não pega as emissões recentes | O merge pega a emissão MAIS RECENTE por matrícula. Se não aparece, verifique no `logs/historico_emissoes.csv` |
| "Falha de autenticação SMTP" | Gmail exige Senha de App, não senha normal |
| PIN esquecido e sem e-mail de recuperação | Apague `config.json` (perde só o PIN, não o histórico) |

---

## Suporte

**Artur Gabriel Oliveira da Silva**
📧 **arturg_oliveira@outlook.com**
