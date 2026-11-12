# -*- coding: utf-8 -*-
"""
Gera uma pagina HTML autonoma com Inadimplentes, links para o WhatsApp Web,
status de contato, campo de observacao e exportacao para uma planilha .xlsx.

Uso no Google Colab:
1. Execute este arquivo ou a celula do notebook fornecido.
2. Selecione a planilha .xls ou .xlsx quando a janela de upload aparecer.
3. Informe o nome do atendente e da empresa (ou dê Enter para os padrões).
4. O HTML sera gerado e baixado automaticamente.
5. Na pagina, use o botao Exportar para Excel (.xlsx) para gerar a planilha sem a coluna do WhatsApp.

Uso local opcional:
    python gerar_inadimplentes_whatsapp_colab.py arquivo.xls saida.html
"""

from __future__ import annotations

import html
import importlib
import json
import os
import re
import subprocess
import sys
import unicodedata
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import quote
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


# Altere somente este valor caso utilize outro DDD para numeros sem DDD.
DDD_PADRAO = "61"

# Nome do arquivo HTML gerado.
NOME_ARQUIVO_SAIDA = "Inadimplentes_cobranca.html"
VERSAO_GERADOR = "2.8"

# Valores padrão caso nada seja digitado no prompt
PADRAO_ATENDENTE = "Nome do atendente"
PADRAO_EMPRESA = "Nome da empresa"

# O Google Colab normalmente opera em UTC. Por isso, a saudacao e definida
# explicitamente com base no horario de Brasilia no momento em que o HTML e criado.
FUSO_HORARIO = "America/Sao_Paulo"
HORA_INICIO_MANHA = 5
HORA_INICIO_TARDE = 12
HORA_INICIO_NOITE = 18

MENSAGEM_SINGULAR = """{saudacao}, tudo bem?
Me chamo {nome_atendente}, do departamento Financeiro da {nome_empresa}.

Referente à mensalidade com vencimento em {datas_vencimento}, no valor de {valor_cobranca}, informamos que, em consulta ao nosso sistema, o valor do boleto consta em aberto.

Poderia informar a previsão para pagamento? Caso já tenha pago, por gentileza, nos encaminhe o comprovante.

Agradecemos a compreensão e permanecemos à disposição para quaisquer esclarecimentos.

Atenciosamente,
Financeiro {nome_empresa}."""

MENSAGEM_PLURAL = """{saudacao}, tudo bem?
Me chamo {nome_atendente}, do departamento Financeiro da {nome_empresa}.

Referente às mensalidades com vencimento em {datas_vencimento}, no valor total de {valor_cobranca}, informamos que, em consulta ao nosso sistema, os valores dos boletos constam em aberto.

Poderia informar a previsão para pagamento? Caso já tenha pago, por gentileza, nos encaminhe os comprovantes.

Agradecemos a compreensão e permanecemos à disposição para quaisquer esclarecimentos.

Atenciosamente,
Financeiro {nome_empresa}."""


# -----------------------------------------------------------------------------
# Instalacao automatica das bibliotecas necessarias
# -----------------------------------------------------------------------------

def garantir_biblioteca(pacote: str, modulo: str | None = None) -> None:
    modulo = modulo or pacote
    try:
        importlib.import_module(modulo)
    except ModuleNotFoundError:
        print(f"Instalando biblioteca necessaria: {pacote}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pacote])


def preparar_bibliotecas(extensao: str) -> None:
    if extensao == ".xls":
        garantir_biblioteca("xlrd==2.0.1", "xlrd")
    elif extensao == ".xlsx":
        garantir_biblioteca("openpyxl>=3.1.0", "openpyxl")
    else:
        raise ValueError("Formato nao suportado. Envie uma planilha .xls ou .xlsx.")


# -----------------------------------------------------------------------------
# Interação via Console / Prompt
# -----------------------------------------------------------------------------

def coletar_identificacao() -> tuple[str, str]:
    """Solicita os dados ao usuário com fallback para os padrões solicitados."""
    print("\n--- Configuração da Cobrança ---")
    atendente = input(f"Informe o seu nome (atendente) [{PADRAO_ATENDENTE}]: ").strip()
    if not atendente:
        atendente = PADRAO_ATENDENTE

    empresa = input(f"Informe o nome da empresa [{PADRAO_EMPRESA}]: ").strip()
    if not empresa:
        empresa = PADRAO_EMPRESA

    print(f"-> Definido: Atendente: '{atendente}' | Empresa: '{empresa}'\n")
    return atendente, empresa


# -----------------------------------------------------------------------------
# Leitura da planilha
# -----------------------------------------------------------------------------

def texto(valor: Any) -> str:
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def normalizar_texto(valor: Any) -> str:
    s = texto(valor)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip().casefold()
    return s


def contem_marcador_cliente(linhas: list[list[Any]]) -> bool:
    """Reconhece os marcadores no relatorio."""
    marcadores = {"informacoes do aluno:", "aluno:", "informacoes do inadimplente:", "inadimplente:"}
    return any(
        any(normalizar_texto(celula) in marcadores for celula in linha)
        for linha in linhas
    )


def ler_xls(caminho: str) -> tuple[list[list[Any]], str]:
    import xlrd

    livro = xlrd.open_workbook(caminho, formatting_info=False)
    melhor_linhas: list[list[Any]] | None = None
    melhor_nome = livro.sheet_names()[0]

    for aba in livro.sheets():
        linhas: list[list[Any]] = []
        for indice_linha in range(aba.nrows):
            linha: list[Any] = []
            for indice_coluna in range(aba.ncols):
                celula = aba.cell(indice_linha, indice_coluna)
                valor = celula.value
                if celula.ctype == xlrd.XL_CELL_DATE:
                    try:
                        valor = xlrd.xldate_as_datetime(valor, livro.datemode)
                    except Exception:
                        pass
                elif isinstance(valor, float) and valor.is_integer():
                    valor = int(valor)
                linha.append(valor)
            linhas.append(linha)

        if melhor_linhas is None:
            melhor_linhas = linhas
        if contem_marcador_cliente(linhas):
            return linhas, aba.name

    return melhor_linhas or [], melhor_nome


def ler_xlsx(caminho: str) -> tuple[list[list[Any]], str]:
    from openpyxl import load_workbook

    livro = load_workbook(caminho, data_only=True, read_only=True)
    melhor_linhas: list[list[Any]] | None = None
    melhor_nome = livro.sheetnames[0]

    for aba in livro.worksheets:
        linhas = [list(linha) for linha in aba.iter_rows(values_only=True)]
        if melhor_linhas is None:
            melhor_linhas = linhas
        if contem_marcador_cliente(linhas):
            livro.close()
            return linhas, aba.title

    livro.close()
    return melhor_linhas or [], melhor_nome


def ler_planilha(caminho: str) -> tuple[list[list[Any]], str]:
    extensao = Path(caminho).suffix.lower()
    preparar_bibliotecas(extensao)
    if extensao == ".xls":
        return ler_xls(caminho)
    return ler_xlsx(caminho)


# -----------------------------------------------------------------------------
# Extracao dos dados do relatorio de Inadimplentes
# -----------------------------------------------------------------------------

ROTULOS_CONHECIDOS = {
    "informacoes do aluno:",
    "informacoes do inadimplente:",
    "informacoes do responsavel:",
    "idade:",
    "endereco:",
    "bairro:",
    "cidade:",
    "cep:",
    "rg:",
    "cpf:",
    "fone:",
    "celular:",
    "fone com.:",
    "responsavel:",
    "categoria",
    "n parc.",
    "no parc.",
    "valor",
    "juros/multa",
    "vlr. c/ jur./multa",
    "vencimento",
    "dias atraso",
    "data negativ.",
    "total vencido:",
    "aluno:",
    "inadimplente:",
    "telefone:",
    "nº parc",
    "no parc",
    "venc.",
    "atraso",
    "totais:",
}


def proximo_valor_nao_vazio(
    linha: list[Any], inicio: int, fim: int | None = None
) -> str:
    limite = len(linha) if fim is None else min(fim, len(linha))
    for indice in range(inicio, limite):
        valor = texto(linha[indice])
        if valor and normalizar_texto(valor) not in ROTULOS_CONHECIDOS:
            return valor
    return ""


def valor_apos_rotulo(
    linha: list[Any], rotulo: str, distancia_maxima: int = 4
) -> str:
    alvo = normalizar_texto(rotulo)
    for indice, valor in enumerate(linha):
        if normalizar_texto(valor) == alvo:
            return proximo_valor_nao_vazio(
                linha, indice + 1, indice + 1 + distancia_maxima
            )
    return ""


def decimal_brasileiro(valor: Any) -> Decimal:
    if valor is None or texto(valor) == "":
        return Decimal("0")
    if isinstance(valor, Decimal):
        return valor
    if isinstance(valor, (int, float)):
        return Decimal(str(valor))

    s = texto(valor).replace("R$", "").replace("\xa0", "").replace(" ", "")
    if s.startswith("="):
        return Decimal("0")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return Decimal("0")


def formatar_reais(valor: Decimal | float | int) -> str:
    if not isinstance(valor, Decimal):
        valor = Decimal(str(valor))
    valor = valor.quantize(Decimal("0.01"))
    sinal = "-" if valor < 0 else ""
    valor = abs(valor)
    inteiro, centavos = f"{valor:.2f}".split(".")
    inteiro = f"{int(inteiro):,}".replace(",", ".")
    return f"{sinal}R$ {inteiro},{centavos}"


def localizar_colunas_cobranca(linha_cabecalho: list[Any]) -> dict[str, int]:
    colunas: dict[str, int] = {}
    for indice, valor in enumerate(linha_cabecalho):
        n = normalizar_texto(valor)
        if n == "categoria":
            colunas["categoria"] = indice
        elif "parc" in n and ("n" in n or "º" in texto(valor)):
            colunas["parcela"] = indice
        elif n == "valor":
            colunas["valor"] = indice
        elif n == "juros/multa":
            colunas["juros"] = indice
        elif ("vlr" in n and "jur" in n) or n == "valor atualizado":
            colunas["atualizado"] = indice
        elif n in {"vencimento", "venc."}:
            colunas["vencimento"] = indice
        elif n in {"dias atraso", "atraso"}:
            colunas["dias_atraso"] = indice
    return colunas


def obter_celula(linha: list[Any], coluna: int | None) -> Any:
    if coluna is None or coluna >= len(linha):
        return ""
    return linha[coluna]


def _indices_inicio_registro(linhas: list[list[Any]]) -> list[int]:
    marcadores = {"informacoes do aluno:", "aluno:", "informacoes do inadimplente:", "inadimplente:"}
    return [
        indice
        for indice, linha in enumerate(linhas)
        if any(normalizar_texto(celula) in marcadores for celula in linha)
    ]


def _extrair_nome_principal(linha: list[Any]) -> str:
    normalizados = [normalizar_texto(v) for v in linha]
    for marcador in ("informacoes do aluno:", "aluno:", "informacoes do inadimplente:", "inadimplente:"):
        if marcador in normalizados:
            i = normalizados.index(marcador)
            return proximo_valor_nao_vazio(linha, i + 1)
    return ""


def extrair_registros(linhas: list[list[Any]]) -> list[dict[str, Any]]:
    inicios = _indices_inicio_registro(linhas)
    registros: list[dict[str, Any]] = []

    for posicao, inicio in enumerate(inicios):
        fim = inicios[posicao + 1] if posicao + 1 < len(inicios) else len(linhas)
        bloco = linhas[inicio:fim]
        nome_principal = _extrair_nome_principal(linhas[inicio])

        responsavel = ""
        celular = ""
        telefone_responsavel = ""
        telefone_principal = ""
        indice_responsavel = len(bloco) + 1
        indice_cabecalho = None
        colunas: dict[str, int] = {}

        for indice_bloco, linha in enumerate(bloco):
            normalizados = [normalizar_texto(valor) for valor in linha]

            if "informacoes do responsavel:" in normalizados:
                indice_responsavel = indice_bloco
                coluna_rotulo = normalizados.index("informacoes do responsavel:")
                responsavel = proximo_valor_nao_vazio(linha, coluna_rotulo + 1)

            if "responsavel:" in normalizados and not responsavel:
                indice_responsavel = indice_bloco
                coluna_rotulo = normalizados.index("responsavel:")
                responsavel = proximo_valor_nao_vazio(linha, coluna_rotulo + 1)
                responsavel = re.sub(r"\s*\([^)]*\)\s*$", "", responsavel).strip()

            if "celular:" in normalizados and not celular:
                celular = valor_apos_rotulo(linha, "Celular:", distancia_maxima=4)

            tem_categoria = "categoria" in normalizados
            tem_vencimento = "vencimento" in normalizados or "venc." in normalizados
            if tem_categoria and tem_vencimento:
                indice_cabecalho = indice_bloco
                colunas = localizar_colunas_cobranca(linha)

        for indice_bloco, linha in enumerate(bloco):
            normalizados = [normalizar_texto(valor) for valor in linha]
            for indice_coluna, n in enumerate(normalizados):
                if n in {"fone:", "telefone:"}:
                    candidato = proximo_valor_nao_vazio(
                        linha,
                        indice_coluna + 1,
                        min(indice_coluna + 5, len(linha)),
                    )
                    if not candidato:
                        continue
                    if indice_bloco < indice_responsavel:
                        telefone_principal = telefone_principal or candidato
                    else:
                        telefone_responsavel = telefone_responsavel or candidato

        telefone = celular or telefone_responsavel or telefone_principal

        parcelas: list[dict[str, Any]] = []
        if indice_cabecalho is not None and colunas:
            for linha in bloco[indice_cabecalho + 1 :]:
                categoria = texto(obter_celula(linha, colunas.get("categoria")))
                categoria_norm = normalizar_texto(categoria)
                if categoria_norm in {"total vencido:", "totais:"}:
                    break
                vencimento = texto(obter_celula(linha, colunas.get("vencimento")))
                if not categoria or not vencimento:
                    continue
                if categoria.startswith("_"):
                    continue

                parcela = texto(obter_celula(linha, colunas.get("parcela")))
                valor_original = decimal_brasileiro(
                    obter_celula(linha, colunas.get("valor"))
                )
                juros = decimal_brasileiro(obter_celula(linha, colunas.get("juros")))

                valor_atualizado = decimal_brasileiro(
                    obter_celula(linha, colunas.get("atualizado"))
                )
                if valor_atualizado == Decimal("0"):
                    valor_atualizado = valor_original + juros

                dias_atraso = texto(obter_celula(linha, colunas.get("dias_atraso")))

                parcelas.append(
                    {
                        "categoria": categoria,
                        "parcela": parcela,
                        "valor_original": valor_original,
                        "juros": juros,
                        "valor_atualizado": valor_atualizado,
                        "vencimento": vencimento,
                        "dias_atraso": dias_atraso,
                    }
                )

        if not nome_principal or not parcelas:
            continue

        total_original = sum(
            (parcela["valor_original"] for parcela in parcelas), Decimal("0")
        )
        total_juros = sum((parcela["juros"] for parcela in parcelas), Decimal("0"))
        total_atualizado = sum(
            (parcela["valor_atualizado"] for parcela in parcelas), Decimal("0")
        )

        registros.append(
            {
                "nome_principal": nome_principal,
                "responsavel": responsavel,
                "telefone_original": telefone,
                "origem_telefone": (
                    "Celular do responsavel"
                    if celular
                    else "Telefone do responsavel"
                    if telefone_responsavel
                    else "Telefone principal"
                    if telefone_principal
                    else "Nao informado"
                ),
                "parcelas": parcelas,
                "total_original": total_original,
                "total_juros": total_juros,
                "total_atualizado": total_atualizado,
            }
        )

    return sorted(registros, key=lambda registro: normalizar_texto(registro["nome_principal"]))


# -----------------------------------------------------------------------------
# WhatsApp e HTML
# -----------------------------------------------------------------------------

def normalizar_telefone_whatsapp(telefone: str, ddd_padrao: str = DDD_PADRAO) -> str:
    bruto = texto(telefone).strip()
    digitos = re.sub(r"\D", "", bruto).lstrip("0")
    pais_explicito = bruto.startswith("+") or re.match(r"^00\s*55", bruto) is not None

    while digitos.startswith("55") and len(digitos) > 13:
        digitos = digitos[2:]

    if digitos.startswith("55") and len(digitos) in (12, 13):
        return digitos

    if pais_explicito and digitos.startswith("55"):
        return ""

    if len(digitos) in (10, 11):
        return "55" + digitos

    if len(digitos) in (8, 9):
        return "55" + ddd_padrao + digitos

    return ""


def juntar_datas(datas: Iterable[str]) -> str:
    unicas: list[str] = []
    for data_vencimento in datas:
        if data_vencimento and data_vencimento not in unicas:
            unicas.append(data_vencimento)
    if not unicas:
        return "data nao identificada"
    if len(unicas) == 1:
        return unicas[0]
    return ", ".join(unicas[:-1]) + " e " + unicas[-1]


def obter_horario_geracao() -> datetime:
    try:
        return datetime.now(ZoneInfo(FUSO_HORARIO))
    except ZoneInfoNotFoundError:
        return datetime.now(timezone(timedelta(hours=-3)))


def definir_saudacao(horario: datetime) -> str:
    hora = horario.hour
    if HORA_INICIO_MANHA <= hora < HORA_INICIO_TARDE:
        return "Bom dia"
    if HORA_INICIO_TARDE <= hora < HORA_INICIO_NOITE:
        return "Boa tarde"
    return "Boa noite"


def criar_link_whatsapp(
    registro: dict[str, Any], saudacao: str, nome_atendente: str, nome_empresa: str
) -> tuple[str, str, str]:
    parcelas = registro["parcelas"]
    datas = juntar_datas(parcela["vencimento"] for parcela in parcelas)

    if len(parcelas) == 1:
        valor_cobranca = formatar_reais(parcelas[0]["valor_atualizado"])
        modelo_mensagem = MENSAGEM_SINGULAR
    else:
        valor_cobranca = formatar_reais(registro["total_atualizado"])
        modelo_mensagem = MENSAGEM_PLURAL

    mensagem = modelo_mensagem.format(
        saudacao=saudacao,
        datas_vencimento=datas,
        valor_cobranca=valor_cobranca,
        nome_atendente=nome_atendente,
        nome_empresa=nome_empresa,
    )
    telefone = normalizar_telefone_whatsapp(registro["telefone_original"])
    if not telefone:
        return "", mensagem, datas
    link = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
    return link, mensagem, datas


def esc(valor: Any) -> str:
    return html.escape(texto(valor), quote=True)


def gerar_script_exportacao_xlsx(dados_exportacao: list[dict[str, Any]], nome_empresa: str) -> str:
    dados_json = json.dumps(
        dados_exportacao,
        ensure_ascii=False,
        separators=(",", ":"),
    ).replace("</", "<\\/")

    modelo = r"""
  <script>
    const DADOS_EXPORTACAO = __DADOS_EXPORTACAO__;
    const CHAVE_ARMAZENAMENTO_OBSERVACOES = 'inadimplentes_observacoes_v1';
    const CHAVE_ARMAZENAMENTO_CONTATADOS = 'inadimplentes_contatados_v1';

    let MEMORIA_OBSERVACOES = {};
    let MEMORIA_CONTATADOS = {};

    function formatarMoedaJs(valor) {
      const v = Number(valor || 0);
      return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    }

    function lerArmazenamento(chaveStorage, memoriaObj) {
      try {
        const conteudo = window.localStorage ? window.localStorage.getItem(chaveStorage) : null;
        if (conteudo) {
          const dados = JSON.parse(conteudo);
          if (dados && typeof dados === 'object') {
            Object.assign(memoriaObj, dados);
          }
        }
      } catch (erro) {
        try {
          const conteudoSession = window.sessionStorage ? window.sessionStorage.getItem(chaveStorage) : null;
          if (conteudoSession) {
            const dados = JSON.parse(conteudoSession);
            if (dados && typeof dados === 'object') {
              Object.assign(memoriaObj, dados);
            }
          }
        } catch (e) {}
      }
      return memoriaObj;
    }

    function salvarArmazenamento(chaveStorage, memoriaObj) {
      try {
        if (window.localStorage) {
          window.localStorage.setItem(chaveStorage, JSON.stringify(memoriaObj));
        }
      } catch (e) {
        try {
          if (window.sessionStorage) {
            window.sessionStorage.setItem(chaveStorage, JSON.stringify(memoriaObj));
          }
        } catch (e2) {}
      }
    }

    function lerObservacoesSalvas() {
      return lerArmazenamento(CHAVE_ARMAZENAMENTO_OBSERVACOES, MEMORIA_OBSERVACOES);
    }

    function lerContatadosSalvos() {
      return lerArmazenamento(CHAVE_ARMAZENAMENTO_CONTATADOS, MEMORIA_CONTATADOS);
    }

    function salvarObservacao(campo) {
      if (!campo || !campo.dataset.chaveObservacao) return;
      const chave = campo.dataset.chaveObservacao;
      const valor = campo.value ? campo.value.trim() : '';

      if (valor) {
        MEMORIA_OBSERVACOES[chave] = valor;
      } else {
        delete MEMORIA_OBSERVACOES[chave];
      }

      salvarArmazenamento(CHAVE_ARMAZENAMENTO_OBSERVACOES, MEMORIA_OBSERVACOES);

      campo.dataset.salvo = 'true';
      const indicador = document.getElementById(`status-obs-${campo.dataset.indice}`);
      if (indicador) {
        indicador.textContent = valor ? '✓ Salvo' : '';
      }
    }

    function alternarContatado(checkbox) {
      if (!checkbox || !checkbox.dataset.chaveContato) return;
      const chave = checkbox.dataset.chaveContato;
      if (checkbox.checked) {
        MEMORIA_CONTATADOS[chave] = true;
      } else {
        delete MEMORIA_CONTATADOS[chave];
      }
      salvarArmazenamento(CHAVE_ARMAZENAMENTO_CONTATADOS, MEMORIA_CONTATADOS);
      atualizarCardsContatados();
    }

    function atualizarCardsContatados() {
      const linhas = document.querySelectorAll('#corpoTabela tr.registro');
      const total = linhas.length;
      let contatados = 0;
      let valorContatados = 0;
      let totalValor = 0;

      linhas.forEach(linha => {
        const cb = linha.querySelector('.flag-contatado');
        const valorRegistro = parseFloat(linha.dataset.valorAtualizado || 0);
        totalValor += valorRegistro;
        if (cb && cb.checked) {
          contatados++;
          valorContatados += valorRegistro;
        }
      });

      const naoContatados = total - contatados;
      const valorNaoContatados = totalValor - valorContatados;

      const pctContatados = total > 0 ? ((contatados / total) * 100).toFixed(1).replace('.', ',') : '0,0';
      const pctNaoContatados = total > 0 ? ((naoContatados / total) * 100).toFixed(1).replace('.', ',') : '0,0';

      const elContatados = document.getElementById('cardContatados');
      const elNaoContatados = document.getElementById('cardNaoContatados');
      if (elContatados) {
        elContatados.innerHTML = `
          <div class="linha-principal">
            <span>${contatados}</span>
            <small class="pct">(${pctContatados}%)</small>
          </div>
          <span class="valor-card">${formatarMoedaJs(valorContatados)}</span>
        `;
      }
      if (elNaoContatados) {
        elNaoContatados.innerHTML = `
          <div class="linha-principal">
            <span>${naoContatados}</span>
            <small class="pct">(${pctNaoContatados}%)</small>
          </div>
          <span class="valor-card">${formatarMoedaJs(valorNaoContatados)}</span>
        `;
      }
    }

    function inicializarObservacoesPersistentes() {
      const observacoes = lerObservacoesSalvas();
      const contatados = lerContatadosSalvos();

      document.querySelectorAll('.campo-observacao').forEach(campo => {
        const chave = campo.dataset.chaveObservacao;
        if (chave && Object.prototype.hasOwnProperty.call(observacoes, chave)) {
          campo.value = observacoes[chave];
          const indicador = document.getElementById(`status-obs-${campo.dataset.indice}`);
          if (indicador && observacoes[chave]) {
            indicador.textContent = '✓ Salvo';
          }
        }

        const sincronizar = () => salvarObservacao(campo);
        campo.addEventListener('input', sincronizar);
        campo.addEventListener('change', sincronizar);
        campo.addEventListener('blur', sincronizar);
      });

      document.querySelectorAll('.flag-contatado').forEach(checkbox => {
        const chave = checkbox.dataset.chaveContato;
        if (chave && contatados[chave]) {
          checkbox.checked = true;
        }
        checkbox.addEventListener('change', () => alternarContatado(checkbox));
      });

      atualizarCardsContatados();
    }

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', inicializarObservacoesPersistentes);
    } else {
      inicializarObservacoesPersistentes();
    }

    function baixarHtmlAtualizado() {
      document.querySelectorAll('.campo-observacao').forEach(salvarObservacao);
      document.querySelectorAll('.campo-observacao').forEach(campo => {
        campo.textContent = campo.value;
      });
      document.querySelectorAll('.flag-contatado').forEach(cb => {
        if (cb.checked) {
          cb.setAttribute('checked', 'checked');
        } else {
          cb.removeAttribute('checked');
        }
      });

      const htmlCompleto = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;
      const blob = new Blob([htmlCompleto], { type: 'text/html;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'Inadimplentes_anotado.html';
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    const TABELA_CRC32 = (() => {
      const tabela = new Uint32Array(256);
      for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) {
          c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        }
        tabela[n] = c >>> 0;
      }
      return tabela;
    })();

    function crc32(bytes) {
      let crc = 0xFFFFFFFF;
      for (let i = 0; i < bytes.length; i++) {
        crc = TABELA_CRC32[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
      }
      return (crc ^ 0xFFFFFFFF) >>> 0;
    }

    function u16(valor) {
      const b = new Uint8Array(2);
      const v = valor >>> 0;
      b[0] = v & 0xFF;
      b[1] = (v >>> 8) & 0xFF;
      return b;
    }

    function u32(valor) {
      const b = new Uint8Array(4);
      const v = valor >>> 0;
      b[0] = v & 0xFF;
      b[1] = (v >>> 8) & 0xFF;
      b[2] = (v >>> 16) & 0xFF;
      b[3] = (v >>> 24) & 0xFF;
      return b;
    }

    function concatenar(partes) {
      const tamanho = partes.reduce((soma, parte) => soma + parte.length, 0);
      const saida = new Uint8Array(tamanho);
      let posicao = 0;
      partes.forEach(parte => {
        saida.set(parte, posicao);
        posicao += parte.length;
      });
      return saida;
    }

    function dataHoraDos() {
      const agora = new Date();
      const ano = Math.max(1980, agora.getFullYear());
      const hora = ((agora.getHours() & 31) << 11)
        | ((agora.getMinutes() & 63) << 5)
        | ((Math.floor(agora.getSeconds() / 2)) & 31);
      const data = (((ano - 1980) & 127) << 9)
        | (((agora.getMonth() + 1) & 15) << 5)
        | (agora.getDate() & 31);
      return { hora, data };
    }

    function criarZipSemCompressao(arquivos) {
      const encoder = new TextEncoder();
      const locais = [];
      const centrais = [];
      const momento = dataHoraDos();
      let deslocamento = 0;

      arquivos.forEach(arquivo => {
        const nome = encoder.encode(arquivo.nome);
        const dados = typeof arquivo.conteudo === 'string'
          ? encoder.encode(arquivo.conteudo)
          : arquivo.conteudo;
        const crc = crc32(dados);

        const cabecalhoLocal = concatenar([
          u32(0x04034B50), u16(20), u16(0x0800), u16(0),
          u16(momento.hora), u16(momento.data), u32(crc),
          u32(dados.length), u32(dados.length), u16(nome.length), u16(0), nome
        ]);
        const registroLocal = concatenar([cabecalhoLocal, dados]);
        locais.push(registroLocal);

        const cabecalhoCentral = concatenar([
          u32(0x02014B50), u16(20), u16(20), u16(0x0800), u16(0),
          u16(momento.hora), u16(momento.data), u32(crc),
          u32(dados.length), u32(dados.length), u16(nome.length),
          u16(0), u16(0), u16(0), u16(0), u32(0),
          u32(deslocamento), nome
        ]);
        centrais.push(cabecalhoCentral);
        deslocamento += registroLocal.length;
      });

      const diretorioCentral = concatenar(centrais);
      const fim = concatenar([
        u32(0x06054B50), u16(0), u16(0),
        u16(arquivos.length), u16(arquivos.length),
        u32(diretorioCentral.length), u32(deslocamento), u16(0)
      ]);
      return concatenar([...locais, diretorioCentral, fim]);
    }

    function escaparXml(valor) {
      return String(valor ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    }

    function colunaExcel(indice) {
      let resultado = '';
      let numero = indice + 1;
      while (numero > 0) {
        const resto = (numero - 1) % 26;
        resultado = String.fromCharCode(65 + resto) + resultado;
        numero = Math.floor((numero - 1) / 26);
      }
      return resultado;
    }

    function celulaXml(linha, coluna, valor, tipo, estilo) {
      const referencia = `${colunaExcel(coluna)}${linha}`;
      if (tipo === 'numero') {
        const numero = Number(valor || 0);
        return `<c r="${referencia}" s="${estilo}"><v>${Number.isFinite(numero) ? numero : 0}</v></c>`;
      }
      const texto = escaparXml(valor);
      return `<c r="${referencia}" s="${estilo}" t="inlineStr"><is><t xml:space="preserve">${texto}</t></is></c>`;
    }

    function montarArquivoXlsx(linhas) {
      const cabecalhos = [
        '#',
        'Nome / Responsável',
        'Telefone',
        'Mensalidades e parcelas em aberto',
        'Valor original',
        'Juros / multa',
        'Total atualizado',
        'Contatado?',
        'Observação'
      ];

      const linhasXml = [];
      linhasXml.push(
        `<row r="1" ht="28" customHeight="1">${cabecalhos.map((titulo, i) => celulaXml(1, i, titulo, 'texto', 1)).join('')}</row>`
      );

      linhas.forEach((linha, indice) => {
        const numeroLinha = indice + 2;
        const valores = [
          [linha.numero, 'numero', 3],
          [linha.nomeResponsavel, 'texto', 3],
          [linha.telefone, 'texto', 3],
          [linha.parcelas, 'texto', 3],
          [linha.valorOriginal, 'numero', 2],
          [linha.jurosMulta, 'numero', 2],
          [linha.totalAtualizado, 'numero', 2],
          [linha.contatado, 'texto', 1],
          [linha.observacao, 'texto', 3]
        ];
        linhasXml.push(
          `<row r="${numeroLinha}">${valores.map((item, i) => celulaXml(numeroLinha, i, item[0], item[1], item[2])).join('')}</row>`
        );
      });

      const ultimaLinha = Math.max(1, linhas.length + 1);
      const planilhaXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:I${ultimaLinha}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="7" customWidth="1"/>
    <col min="2" max="2" width="38" customWidth="1"/>
    <col min="3" max="3" width="27" customWidth="1"/>
    <col min="4" max="4" width="75" customWidth="1"/>
    <col min="5" max="7" width="18" customWidth="1"/>
    <col min="8" max="8" width="16" customWidth="1"/>
    <col min="9" max="9" width="42" customWidth="1"/>
  </cols>
  <sheetData>${linhasXml.join('')}</sheetData>
  <autoFilter ref="A1:I${ultimaLinha}"/>
  <pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>
  <pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;

      const estilosXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;R$&quot; #,##0.00"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF143A66"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD7DEE8"/></left>
      <right style="thin"><color rgb="FFD7DEE8"/></right>
      <top style="thin"><color rgb="FFD7DEE8"/></top>
      <bottom style="thin"><color rgb="FFD7DEE8"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

      const dataIso = new Date().toISOString();
      const arquivos = [
        {
          nome: '[Content_Types].xml',
          conteudo: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
        },
        {
          nome: '_rels/.rels',
          conteudo: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
        },
        {
          nome: 'xl/workbook.xml',
          conteudo: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Inadimplentes" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
        },
        {
          nome: 'xl/_rels/workbook.xml.rels',
          conteudo: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
        },
        { nome: 'xl/worksheets/sheet1.xml', conteudo: planilhaXml },
        { nome: 'xl/styles.xml', conteudo: estilosXml },
        {
          nome: 'docProps/core.xml',
          conteudo: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Inadimplentes - __NOME_EMPRESA__</dc:title>
  <dc:creator>Gerador Automático</dc:creator>
  <cp:lastModifiedBy>Gerador Automático</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${dataIso}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${dataIso}</dcterms:modified>
</cp:coreProperties>`
        },
        {
          nome: 'docProps/app.xml',
          conteudo: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Gerador Automático</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Planilhas</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Inadimplentes</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>`
        }
      ];

      return criarZipSemCompressao(arquivos);
    }

    function exportarXLSX() {
      try {
        document.querySelectorAll('.campo-observacao').forEach(salvarObservacao);
        const observacoes = lerObservacoesSalvas();
        const contatados = lerContatadosSalvos();

        const linhas = DADOS_EXPORTACAO.map(item => {
          const campo = document.getElementById(`observacao-${item.numero}`);
          const valorCampo = campo ? campo.value.trim() : '';
          const valorMemoria = observacoes[item.chaveObservacao] || '';

          const checkbox = document.getElementById(`contatado-${item.numero}`);
          const estaContatado = checkbox ? checkbox.checked : Boolean(contatados[item.chaveObservacao]);

          return {
            numero: item.numero,
            nomeResponsavel: `${item.nomePrincipal}\nResponsável: ${item.responsavel || 'Não informado'}`,
            telefone: `${item.telefone || 'Não informado'}${item.origemTelefone ? `\n${item.origemTelefone}` : ''}`,
            parcelas: item.parcelas,
            valorOriginal: item.valorOriginal,
            jurosMulta: item.jurosMulta,
            totalAtualizado: item.totalAtualizado,
            contatado: estaContatado ? 'Sim' : 'Não',
            observacao: valorCampo || valorMemoria
          };
        });

        const bytes = montarArquivoXlsx(linhas);
        const blob = new Blob(
          [bytes],
          { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
        );
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'Inadimplentes_exportado.xlsx';
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      } catch (erro) {
        console.error(erro);
        alert('Não foi possível gerar a planilha. Recarregue a página e tente novamente.');
      }
    }
  </script>
"""
    return modelo.replace("__DADOS_EXPORTACAO__", dados_json).replace("__NOME_EMPRESA__", esc(nome_empresa))


def gerar_html(
    registros: list[dict[str, Any]],
    nome_aba: str,
    nome_atendente: str = PADRAO_ATENDENTE,
    nome_empresa: str = PADRAO_EMPRESA,
) -> str:
    horario_geracao = obter_horario_geracao()
    saudacao = definir_saudacao(horario_geracao)
    horario_geracao_formatado = horario_geracao.strftime("%d/%m/%Y às %H:%M")

    linhas_html: list[str] = []
    dados_exportacao: list[dict[str, Any]] = []
    total_geral = sum(
        (registro["total_atualizado"] for registro in registros), Decimal("0")
    )
    com_whatsapp = 0

    for indice, registro in enumerate(registros, start=1):
        link, mensagem, datas = criar_link_whatsapp(
            registro, saudacao, nome_atendente=nome_atendente, nome_empresa=nome_empresa
        )
        if link:
            com_whatsapp += 1

        itens_parcelas = []
        itens_parcelas_exportacao: list[str] = []
        for parcela in registro["parcelas"]:
            numero_parcela = (
                f" - parcela {esc(parcela['parcela'])}" if parcela["parcela"] else ""
            )
            atraso = (
                f" - {esc(parcela['dias_atraso'])} dia(s) em atraso"
                if parcela["dias_atraso"]
                else ""
            )
            itens_parcelas.append(
                "<li>"
                f"<strong>{esc(parcela['categoria'])}</strong>{numero_parcela}"
                f"<br><span>Vencimento: {esc(parcela['vencimento'])}{atraso}</span>"
                f"<br><span>Original: {formatar_reais(parcela['valor_original'])} | "
                f"Juros/multa: {formatar_reais(parcela['juros'])} | "
                f"Atualizado: <strong>{formatar_reais(parcela['valor_atualizado'])}</strong></span>"
                "</li>"
            )

            parcela_exportacao = parcela["categoria"]
            if parcela["parcela"]:
                parcela_exportacao += f" - parcela {parcela['parcela']}"
            parcela_exportacao += f" | Vencimento: {parcela['vencimento']}"
            if parcela["dias_atraso"]:
                parcela_exportacao += f" | {parcela['dias_atraso']} dia(s) em atraso"
            parcela_exportacao += (
                f" | Original: {formatar_reais(parcela['valor_original'])}"
                f" | Juros/multa: {formatar_reais(parcela['juros'])}"
                f" | Atualizado: {formatar_reais(parcela['valor_atualizado'])}"
            )
            itens_parcelas_exportacao.append(parcela_exportacao)

        if link:
            botao = (
                f'<a class="btn-whatsapp" href="{esc(link)}" target="_blank" '
                'rel="noopener noreferrer">Abrir WhatsApp</a>'
            )
            status_whatsapp = "com-whatsapp"
        else:
            botao = '<span class="sem-telefone">Telefone ausente ou invalido</span>'
            status_whatsapp = "sem-whatsapp"

        telefone_exibido = esc(registro["telefone_original"]) or "Nao informado"
        texto_busca = " ".join(
            [
                registro["nome_principal"],
                registro["responsavel"],
                registro["telefone_original"],
                datas,
                " ".join(p["categoria"] for p in registro["parcelas"]),
            ]
        )

        chave_obs = normalizar_texto(
            f"{registro['nome_principal']}|{registro['responsavel']}|{registro['telefone_original']}"
        )

        dados_exportacao.append(
            {
                "numero": indice,
                "nomePrincipal": registro["nome_principal"],
                "responsavel": registro["responsavel"],
                "telefone": registro["telefone_original"],
                "origemTelefone": registro["origem_telefone"],
                "parcelas": "\n".join(itens_parcelas_exportacao),
                "valorOriginal": float(registro["total_original"]),
                "jurosMulta": float(registro["total_juros"]),
                "totalAtualizado": float(registro["total_atualizado"]),
                "chaveObservacao": chave_obs,
            }
        )

        valor_atualizado_flt = float(registro["total_atualizado"])

        linhas_html.append(
            f"""
            <tr class="registro {status_whatsapp}"
                data-search="{esc(normalizar_texto(texto_busca))}"
                data-valor-atualizado="{valor_atualizado_flt:.2f}">
              <td class="numero">{indice}</td>
              <td>
                <strong class="nome-principal">{esc(registro['nome_principal'])}</strong>
                <span class="secundario">Responsavel: {esc(registro['responsavel']) or 'Nao informado'}</span>
              </td>
              <td>
                <strong>{telefone_exibido}</strong>
                <span class="secundario">{esc(registro['origem_telefone'])}</span>
              </td>
              <td><ul class="lista-parcelas">{''.join(itens_parcelas)}</ul></td>
              <td class="valor">{formatar_reais(registro['total_original'])}</td>
              <td class="valor">{formatar_reais(registro['total_juros'])}</td>
              <td class="valor total">{formatar_reais(registro['total_atualizado'])}</td>
              <td class="acao coluna-whatsapp">
                {botao}
                <details>
                  <summary>Ver mensagem</summary>
                  <pre>{esc(mensagem)}</pre>
                </details>
              </td>
              <td class="coluna-contatado">
                <label class="container-flag">
                  <input type="checkbox" id="contatado-{indice}" class="flag-contatado"
                    data-indice="{indice}" data-chave-contato="{esc(chave_obs)}">
                  <span class="flag-customizada"></span>
                </label>
              </td>
              <td class="observacao">
                <textarea id="observacao-{indice}" class="campo-observacao"
                  data-indice="{indice}"
                  data-chave-observacao="{esc(chave_obs)}"
                  placeholder="Digite uma observação..."
                  aria-label="Observação de {esc(registro['nome_principal'])}"></textarea>
                <span id="status-obs-{indice}" class="status-obs"></span>
              </td>
            </tr>
            """
        )

    sem_whatsapp = len(registros) - com_whatsapp
    script_exportacao = gerar_script_exportacao_xlsx(dados_exportacao, nome_empresa)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cobrança de Inadimplentes - {esc(nome_empresa)}</title>
  <style>
    :root {{
      --azul: #143a66;
      --azul-claro: #eaf2fb;
      --verde: #128c4a;
      --verde-hover: #0d733c;
      --vermelho: #a52a2a;
      --cinza: #667085;
      --borda: #d7dee8;
      --fundo: #f4f7fb;
      --branco: #ffffff;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      color: #172033;
      background: var(--fundo);
    }}
    .container {{
      width: min(2000px, 96vw);
      margin: 28px auto;
    }}
    .cabecalho {{
      display: flex;
      justify-content: space-between;
      gap: 20px;
      align-items: flex-end;
      margin-bottom: 18px;
    }}
    h1 {{ margin: 0 0 6px; color: var(--azul); font-size: 28px; }}
    .subtitulo {{ margin: 0; color: var(--cinza); }}
    .acoes-topo {{ display: flex; gap: 10px; flex-wrap: wrap; }}
    button {{
      border: 1px solid var(--borda);
      background: var(--branco);
      padding: 10px 14px;
      border-radius: 8px;
      cursor: pointer;
      font-weight: 700;
    }}
    button:hover {{ background: #f1f4f8; }}
    .btn-exportar {{
      background: var(--azul);
      border-color: var(--azul);
      color: var(--branco);
    }}
    .btn-exportar:hover {{ background: #0d2d52; }}
    .btn-salvar-html {{
      background: #2a6f97;
      border-color: #2a6f97;
      color: var(--branco);
    }}
    .btn-salvar-html:hover {{ background: #1f5270; }}
    .cards {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }}
    .card {{
      background: var(--branco);
      border: 1px solid var(--borda);
      border-radius: 12px;
      padding: 16px;
      box-shadow: 0 3px 12px rgba(20, 58, 102, .05);
    }}
    .card span {{ color: var(--cinza); font-size: 13px; display: block; margin-bottom: 8px; }}
    .card strong {{
      font-size: 24px;
      color: var(--azul);
      display: flex;
      flex-direction: column;
      gap: 4px;
    }}
    .card .linha-principal {{
      display: flex;
      align-items: baseline;
      gap: 6px;
      flex-wrap: wrap;
    }}
    .card strong small.pct {{
      font-size: 14px;
      font-weight: 600;
      color: var(--cinza);
    }}
    .card strong .valor-card {{
      font-size: 15px;
      font-weight: 700;
      color: var(--azul);
    }}
    .filtros {{
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
      padding: 14px;
      background: var(--branco);
      border: 1px solid var(--borda);
      border-radius: 12px 12px 0 0;
    }}
    .filtros input[type="search"] {{
      flex: 1 1 360px;
      min-width: 220px;
      padding: 11px 13px;
      border: 1px solid var(--borda);
      border-radius: 8px;
      font-size: 15px;
    }}
    .filtros label {{ color: var(--cinza); font-size: 14px; }}
    .contador {{ margin-left: auto; color: var(--cinza); font-weight: 700; }}
    .tabela-wrapper {{
      overflow-x: auto;
      background: var(--branco);
      border: 1px solid var(--borda);
      border-top: 0;
      border-radius: 0 0 12px 12px;
    }}
    table {{ width: 100%; border-collapse: collapse; min-width: 1800px; }}
    thead th {{
      position: sticky;
      top: 0;
      z-index: 2;
      background: var(--azul);
      color: var(--branco);
      text-align: left;
      padding: 12px 10px;
      font-size: 13px;
      white-space: nowrap;
    }}
    tbody td {{
      border-top: 1px solid var(--borda);
      padding: 12px 10px;
      vertical-align: top;
      font-size: 13px;
    }}
      th:nth-child(4), td:nth-child(4) {{
        filter: blur(7px);
        user-select: none;
        pointer-events: none;
      }}
    tbody tr:hover {{ background: #f9fbfe; }}
    .numero {{ width: 46px; text-align: center; color: var(--cinza); }}
    .nome-principal {{ font-size: 14px; color: var(--azul); }}
    .secundario {{ display: block; color: var(--cinza); margin-top: 6px; line-height: 1.35; }}
    .lista-parcelas {{ margin: 0; padding-left: 18px; min-width: 390px; }}
    .lista-parcelas li {{ margin-bottom: 9px; line-height: 1.4; }}
    .lista-parcelas li:last-child {{ margin-bottom: 0; }}
    .lista-parcelas span {{ color: #475467; }}
    .valor {{ white-space: nowrap; text-align: right; }}
    .valor.total {{ font-weight: 800; color: var(--vermelho); }}
    .acao {{ min-width: 180px; text-align: center; }}
    .btn-whatsapp {{
      display: inline-block;
      padding: 10px 14px;
      color: var(--branco);
      background: var(--verde);
      border-radius: 8px;
      text-decoration: none;
      font-weight: 800;
      white-space: nowrap;
    }}
    .btn-whatsapp:hover {{ background: var(--verde-hover); }}
    .sem-telefone {{
      display: inline-block;
      padding: 9px;
      color: var(--vermelho);
      background: #fff0f0;
      border: 1px solid #f3caca;
      border-radius: 8px;
      font-weight: 700;
    }}
    .coluna-contatado {{
      width: 90px;
      min-width: 90px;
      text-align: center;
      vertical-align: middle;
    }}
    .container-flag {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      position: relative;
    }}
    .container-flag input[type="checkbox"] {{
      width: 22px;
      height: 22px;
      cursor: pointer;
      accent-color: var(--verde);
    }}
    .observacao {{
      min-width: 270px;
      width: 270px;
      position: sticky;
      right: 0;
      z-index: 1;
      background: var(--branco);
      box-shadow: -3px 0 7px rgba(20, 58, 102, .10);
    }}
    .cabecalho-observacao {{
      position: sticky;
      right: 0;
      z-index: 5;
      min-width: 270px;
      width: 270px;
      box-shadow: -3px 0 7px rgba(20, 58, 102, .16);
    }}
    .coluna-whatsapp {{
      background: var(--branco);
    }}
    .cabecalho-whatsapp {{
      background: var(--azul);
    }}
    tbody tr:hover .observacao {{ background: #f9fbfe; }}
    .campo-observacao {{
      width: 100%;
      min-height: 84px;
      resize: vertical;
      border: 1px solid var(--borda);
      border-radius: 8px;
      padding: 10px;
      font-family: Arial, Helvetica, sans-serif;
      font-size: 13px;
      line-height: 1.4;
      color: #172033;
      background: #fff;
      display: block;
    }}
    .campo-observacao:focus {{
      outline: 2px solid #b9d2ee;
      border-color: var(--azul);
    }}
    .status-obs {{
      display: block;
      font-size: 11px;
      color: var(--verde);
      font-weight: 700;
      margin-top: 4px;
      text-align: right;
      min-height: 14px;
    }}
    details {{ margin-top: 9px; text-align: left; }}
    summary {{ cursor: pointer; color: var(--azul); font-weight: 700; text-align: center; }}
    pre {{
      white-space: pre-wrap;
      overflow-wrap: anywhere;
      background: var(--azul-claro);
      border-radius: 8px;
      padding: 10px;
      font-family: Arial, Helvetica, sans-serif;
      font-size: 12px;
      line-height: 1.45;
      min-width: 300px;
    }}
    .oculto {{ display: none; }}
    .rodape {{ margin: 14px 2px 0; color: var(--cinza); font-size: 12px; }}
    @media (max-width: 900px) {{
      .cabecalho {{ align-items: flex-start; flex-direction: column; }}
      .contador {{ width: 100%; margin-left: 0; }}
    }}
    @media print {{
      body {{ background: white; }}
      .container {{ width: 100%; margin: 0; }}
      .acoes-topo, .filtros, .acao details {{ display: none !important; }}
      .tabela-wrapper {{ overflow: visible; border: 0; }}
      table {{ min-width: 0; font-size: 9px; }}
      thead th {{ position: static; }}
      .btn-whatsapp, .sem-telefone {{ color: black; background: transparent; border: 0; padding: 0; }}
      .campo-observacao {{ border: 0; min-height: 50px; padding: 0; resize: none; }}
      .status-obs {{ display: none; }}
      .observacao, .coluna-whatsapp, .cabecalho-observacao, .cabecalho-whatsapp {{
        position: static; box-shadow: none;
      }}
    }}
  </style>
</head>
<body>
  <main class="container">
    <section class="cabecalho">
      <div>
        <h1>Inadimplentes - {esc(nome_empresa)}</h1>
        <p class="subtitulo">Atendente: {esc(nome_atendente)} | Relacao consolidada | Aba de origem: {esc(nome_aba)} | Versão do gerador: {esc(VERSAO_GERADOR)}</p>
      </div>
      <div class="acoes-topo">
        <button type="button" class="btn-exportar" onclick="exportarXLSX()">Exportar para Excel (.xlsx)</button>
        <button type="button" class="btn-salvar-html" onclick="baixarHtmlAtualizado()">Salvar HTML com Anotações</button>
        <button type="button" onclick="window.print()">Imprimir / Salvar em PDF</button>
        <button type="button" onclick="limparFiltros()">Limpar filtros</button>
      </div>
    </section>

    <section class="cards">
      <div class="card">
        <span>Total de Inadimplentes</span>
        <strong>
          <div class="linha-principal"><span>{len(registros)}</span></div>
          <span class="valor-card">{formatar_reais(total_geral)}</span>
        </strong>
      </div>
      <div class="card">
        <span>Total de Inadimplentes contatados</span>
        <strong id="cardContatados">
          <div class="linha-principal">
            <span>0</span>
            <small class="pct">(0,0%)</small>
          </div>
          <span class="valor-card">R$ 0,00</span>
        </strong>
      </div>
      <div class="card">
        <span>Total de Inadimplentes não contatados</span>
        <strong id="cardNaoContatados">
          <div class="linha-principal">
            <span>{len(registros)}</span>
            <small class="pct">(100,0%)</small>
          </div>
          <span class="valor-card">{formatar_reais(total_geral)}</span>
        </strong>
      </div>
      <div class="card">
        <span>Com link de WhatsApp</span>
        <strong>
          <div class="linha-principal"><span>{com_whatsapp}</span></div>
        </strong>
      </div>
      <div class="card">
        <span>Sem telefone valido</span>
        <strong>
          <div class="linha-principal"><span>{sem_whatsapp}</span></div>
        </strong>
      </div>
      <div class="card">
        <span>Total atualizado em aberto</span>
        <strong>
          <div class="linha-principal"><span>{formatar_reais(total_geral)}</span></div>
        </strong>
      </div>
    </section>

    <section class="filtros">
      <input id="pesquisa" type="search" placeholder="Pesquisar nome, responsavel, telefone, parcela ou vencimento..." oninput="filtrar()">
      <label><input id="somenteWhatsapp" type="checkbox" onchange="filtrar()"> Mostrar somente quem tem WhatsApp</label>
      <label><input id="somenteSemTelefone" type="checkbox" onchange="filtrar()"> Mostrar somente sem telefone</label>
      <label><input id="somenteContatados" type="checkbox" onchange="filtrar()"> Mostrar somente contatados</label>
      <label><input id="somenteNaoContatados" type="checkbox" onchange="filtrar()"> Mostrar somente não contatados</label>
      <span id="contador" class="contador">{len(registros)} registro(s)</span>
    </section>

    <section class="tabela-wrapper">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Nome / Responsavel</th>
            <th>Telefone</th>
            <th>Mensalidades e parcelas em aberto</th>
            <th>Valor original</th>
            <th>Juros / multa</th>
            <th>Total atualizado</th>
            <th class="coluna-whatsapp cabecalho-whatsapp">WhatsApp Web</th>
            <th class="coluna-contatado">Contatado?</th>
            <th class="cabecalho-observacao">Observação</th>
          </tr>
        </thead>
        <tbody id="corpoTabela">
          {''.join(linhas_html)}
        </tbody>
      </table>
    </section>

    <p class="rodape">
      Os links usam o WhatsApp Web. Numeros locais sem DDD receberam automaticamente o DDD {esc(DDD_PADRAO)}.
      A saudacao das mensagens foi definida como &quot;{esc(saudacao)}&quot; no momento da geracao do arquivo,
      em {esc(horario_geracao_formatado)} (horario de Brasilia).
      A exportacao para Excel inclui todos os registros, status de contato e todas as observacoes digitadas.
    </p>
  </main>

  {script_exportacao}
  <script>
    function filtrar() {{
      const termo = document.getElementById('pesquisa').value
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
      const somenteWhatsapp = document.getElementById('somenteWhatsapp').checked;
      const somenteSemTelefone = document.getElementById('somenteSemTelefone').checked;
      const somenteContatados = document.getElementById('somenteContatados').checked;
      const somenteNaoContatados = document.getElementById('somenteNaoContatados').checked;
      const linhas = Array.from(document.querySelectorAll('#corpoTabela tr.registro'));
      let visiveis = 0;

      if (somenteWhatsapp && somenteSemTelefone) {{
        document.getElementById('somenteSemTelefone').checked = false;
      }}
      if (somenteContatados && somenteNaoContatados) {{
        document.getElementById('somenteNaoContatados').checked = false;
      }}

      linhas.forEach(linha => {{
        const correspondeTexto = !termo || linha.dataset.search.includes(termo);
        const temWhatsapp = linha.classList.contains('com-whatsapp');
        const checkbox = linha.querySelector('.flag-contatado');
        const estaContatado = checkbox ? checkbox.checked : false;

        const filtroWhatsapp = !somenteWhatsapp || temWhatsapp;
        const filtroSemTelefone = !document.getElementById('somenteSemTelefone').checked || !temWhatsapp;
        const filtroContatado = !document.getElementById('somenteContatados').checked || estaContatado;
        const filtroNaoContatado = !document.getElementById('somenteNaoContatados').checked || !estaContatado;

        const mostrar = correspondeTexto && filtroWhatsapp && filtroSemTelefone && filtroContatado && filtroNaoContatado;
        linha.classList.toggle('oculto', !mostrar);
        if (mostrar) visiveis++;
      }});

      document.getElementById('contador').textContent = `${{visiveis}} registro(s)`;
    }}

    function limparFiltros() {{
      document.getElementById('pesquisa').value = '';
      document.getElementById('somenteWhatsapp').checked = false;
      document.getElementById('somenteSemTelefone').checked = false;
      document.getElementById('somenteContatados').checked = false;
      document.getElementById('somenteNaoContatados').checked = false;
      filtrar();
    }}
  </script>
</body>
</html>
"""


def salvar_html(conteudo: str, caminho_saida: str) -> str:
    Path(caminho_saida).write_text(conteudo, encoding="utf-8")
    return str(Path(caminho_saida).resolve())


# -----------------------------------------------------------------------------
# Upload e execucao no Google Colab
# -----------------------------------------------------------------------------

def esta_no_google_colab() -> bool:
    try:
        import google.colab  # type: ignore  # noqa: F401
        return True
    except ImportError:
        return False


def selecionar_arquivo_colab() -> str:
    try:
        from google.colab import files
        from IPython.display import Markdown, display
    except ImportError as erro:
        raise RuntimeError(
            "Nenhum arquivo foi informado. Execute no Google Colab ou passe o caminho "
            "da planilha como primeiro argumento."
        ) from erro

    display(
        Markdown(
            "## 📤 Envie a planilha de Inadimplentes\n"
            "Clique no botão **Escolher arquivos** que aparecerá abaixo e selecione "
            "a planilha no formato **.xls** ou **.xlsx**.\n\n"
            "O processamento começará automaticamente após o envio."
        )
    )
    print("Aguardando o envio da planilha...")

    enviados = files.upload()
    if not enviados:
        raise RuntimeError("Nenhum arquivo foi enviado.")

    planilhas_validas = [
        nome for nome in enviados
        if Path(nome).suffix.lower() in {".xls", ".xlsx"}
    ]
    if not planilhas_validas:
        raise ValueError(
            "O arquivo enviado não é uma planilha .xls ou .xlsx. "
            "Execute a célula novamente e envie o arquivo correto."
        )

    if len(planilhas_validas) > 1:
        print(
            f"Foram enviadas {len(planilhas_validas)} planilhas. "
            f"Será utilizada: {planilhas_validas[0]}"
        )

    arquivo_escolhido = planilhas_validas[0]
    print(f"Planilha recebida: {arquivo_escolhido}")
    return arquivo_escolhido


def tentar_baixar_no_colab(caminho_saida: str) -> None:
    try:
        from google.colab import files
        files.download(caminho_saida)
    except ImportError:
        pass


def executar(caminho_entrada: str, caminho_saida: str = NOME_ARQUIVO_SAIDA) -> str:
    print(f"Lendo planilha: {caminho_entrada}")
    linhas, nome_aba = ler_planilha(caminho_entrada)
    registros = extrair_registros(linhas)

    if not registros:
        raise RuntimeError(
            "Nao foi possivel localizar os blocos de Inadimplentes. "
            "Confirme se a planilha possui o formato correto de cobrança."
        )

    nome_atendente, nome_empresa = coletar_identificacao()

    conteudo = gerar_html(registros, nome_aba, nome_atendente=nome_atendente, nome_empresa=nome_empresa)

    elementos_obrigatorios = {
        "coluna Observação": '<th class="cabecalho-observacao">Observação</th>',
        "campos de Observação": 'class="campo-observacao"',
        "botão Exportar para Excel": 'onclick="exportarXLSX()"',
        "função de exportação": 'function exportarXLSX()',
        "persistência das observações": 'function inicializarObservacoesPersistentes()',
        "salvamento automático das observações": 'function salvarObservacao(campo)',
    }
    ausentes = [nome for nome, trecho in elementos_obrigatorios.items() if trecho not in conteudo]
    if ausentes:
        raise RuntimeError(
            "Falha de validação do HTML. Não foram encontrados: " + ", ".join(ausentes)
        )

    caminho_final = salvar_html(conteudo, caminho_saida)

    validos = sum(
        bool(normalizar_telefone_whatsapp(registro["telefone_original"]))
        for registro in registros
    )
    print("\nConcluido com sucesso.")
    print(f"Total de Inadimplentes: {len(registros)}")
    print(f"Links de WhatsApp gerados: {validos}")
    print(f"Registros sem telefone valido: {len(registros) - validos}")
    print(f"Versão do gerador: {VERSAO_GERADOR}")
    print(f"Arquivo HTML: {caminho_final}")
    return caminho_final


def main() -> None:
    if esta_no_google_colab():
        entrada = selecionar_arquivo_colab()
        saida = NOME_ARQUIVO_SAIDA
    elif len(sys.argv) >= 2 and Path(sys.argv[1]).exists():
        entrada = sys.argv[1]
        saida = sys.argv[2] if len(sys.argv) >= 3 else NOME_ARQUIVO_SAIDA
    else:
        raise RuntimeError(
            "Uso local: python gerar_inadimplentes_whatsapp_colab.py "
            "caminho_da_planilha.xls [saida.html]"
        )

    caminho_final = executar(entrada, saida)
    tentar_baixar_no_colab(caminho_final)


if __name__ == "__main__":
    main()
