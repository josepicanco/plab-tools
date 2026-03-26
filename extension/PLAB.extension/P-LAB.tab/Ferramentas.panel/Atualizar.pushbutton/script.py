# -*- coding: utf-8 -*-
"""Verifica e instala atualizacoes dos plugins P-LAB.

Conecta ao repositorio GitHub e compara a versao instalada
com a versao mais recente disponivel.
"""
__title__ = "Atualizar\nP-LAB"
__author__ = "P-LAB Engenharia"
__doc__ = "Verifica atualizacoes dos plugins P-LAB e instala automaticamente."

import os
import sys
import json
import shutil
import tempfile
import zipfile

# Adiciona Python 3 ao path para usar requests
PYTHON3_PATH = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Python", "Python314", "Lib", "site-packages")
if PYTHON3_PATH not in sys.path and os.path.exists(PYTHON3_PATH):
    sys.path.insert(0, PYTHON3_PATH)

# pyRevit imports
from pyrevit import forms, script

logger = script.get_logger()

# ── CONFIGURACAO ─────────────────────────────────────────────────────────────
# SUBSTITUA pelo link do seu repositorio quando criar
GITHUB_USER    = "SEU_USUARIO_GITHUB"
GITHUB_REPO    = "plab-tools"
GITHUB_BRANCH  = "main"

# URLs base
RAW_BASE  = "https://raw.githubusercontent.com/{}/{}/{}".format(GITHUB_USER, GITHUB_REPO, GITHUB_BRANCH)
ZIP_URL   = "https://github.com/{}/{}/archive/refs/heads/{}.zip".format(GITHUB_USER, GITHUB_REPO, GITHUB_BRANCH)

# Pasta onde a extensao esta instalada (dois niveis acima deste script)
# Estrutura: PLAB.extension / P-LAB.tab / Ferramentas.panel / Atualizar.pushbutton / script.py
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
PUSHBUTTON_DIR = SCRIPT_DIR
PANEL_DIR      = os.path.dirname(PUSHBUTTON_DIR)
TAB_DIR        = os.path.dirname(PANEL_DIR)
EXTENSION_DIR  = os.path.dirname(TAB_DIR)

VERSION_FILE_LOCAL  = os.path.join(EXTENSION_DIR, "..", "version.json")
VERSION_FILE_LOCAL  = os.path.normpath(VERSION_FILE_LOCAL)

# ── FUNCOES AUXILIARES ────────────────────────────────────────────────────────

def ler_versao_local():
    """Le o version.json instalado localmente."""
    try:
        if os.path.exists(VERSION_FILE_LOCAL):
            with open(VERSION_FILE_LOCAL, "r") as f:
                return json.load(f)
    except Exception as e:
        logger.warning("Nao foi possivel ler versao local: {}".format(e))
    return {"version": "0.0.0"}


def ler_versao_remota():
    """Baixa o version.json do GitHub e retorna como dict."""
    try:
        # Importa urllib que funciona tanto em IronPython quanto Python 3
        try:
            from urllib2 import urlopen, URLError  # IronPython / Python 2
        except ImportError:
            from urllib.request import urlopen    # Python 3
            from urllib.error import URLError

        url = "{}/version.json".format(RAW_BASE)
        resposta = urlopen(url, timeout=10)
        conteudo = resposta.read()
        if isinstance(conteudo, bytes):
            conteudo = conteudo.decode("utf-8")
        return json.loads(conteudo)

    except Exception as e:
        logger.error("Erro ao conectar ao GitHub: {}".format(e))
        return None


def versao_maior(remota, local):
    """Compara versoes no formato 'X.Y.Z'. Retorna True se remota > local."""
    try:
        def partes(v):
            return [int(x) for x in v.split(".")]
        return partes(remota) > partes(local)
    except Exception:
        return False


def baixar_e_extrair_zip(destino_temp):
    """Baixa o ZIP do repositorio e extrai para pasta temporaria."""
    try:
        try:
            from urllib2 import urlopen
        except ImportError:
            from urllib.request import urlopen

        forms.show_balloon("P-LAB Atualizar", "Baixando atualizacao...")

        resposta  = urlopen(ZIP_URL, timeout=60)
        zip_path  = os.path.join(destino_temp, "plab-tools.zip")

        with open(zip_path, "wb") as f:
            f.write(resposta.read())

        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(destino_temp)

        # O ZIP extrai para uma pasta com o nome "plab-tools-main"
        pasta_extraida = os.path.join(destino_temp, "{}-{}".format(GITHUB_REPO, GITHUB_BRANCH))
        return pasta_extraida

    except Exception as e:
        logger.error("Erro ao baixar ZIP: {}".format(e))
        return None


def copiar_atualizacao(pasta_nova):
    """Substitui os arquivos da extensao com os novos baixados do GitHub."""
    try:
        # Pasta da extensao no ZIP: plab-tools-main/PLAB.extension/
        nova_extension = os.path.join(pasta_nova, "PLAB.extension")
        if not os.path.exists(nova_extension):
            logger.error("Estrutura do repositorio incorreta: pasta PLAB.extension nao encontrada.")
            return False

        # Copia cada arquivo novo para a extensao instalada
        for raiz, pastas, arquivos in os.walk(nova_extension):
            # Caminho relativo dentro da extensao
            relativo = os.path.relpath(raiz, nova_extension)
            destino_dir = os.path.join(EXTENSION_DIR, relativo)

            if not os.path.exists(destino_dir):
                os.makedirs(destino_dir)

            for arquivo in arquivos:
                origem  = os.path.join(raiz, arquivo)
                destino = os.path.join(destino_dir, arquivo)
                shutil.copy2(origem, destino)

        # Atualiza o version.json
        novo_version = os.path.join(pasta_nova, "version.json")
        if os.path.exists(novo_version):
            shutil.copy2(novo_version, VERSION_FILE_LOCAL)

        return True

    except Exception as e:
        logger.error("Erro ao copiar arquivos: {}".format(e))
        return False


# ── EXECUCAO PRINCIPAL ────────────────────────────────────────────────────────

def main():
    # 1. Le versoes
    forms.show_balloon("P-LAB Atualizar", "Verificando atualizacoes...")

    versao_local  = ler_versao_local()
    versao_remota = ler_versao_remota()

    if versao_remota is None:
        forms.alert(
            "Nao foi possivel conectar ao servidor.\n"
            "Verifique sua conexao com a internet e tente novamente.",
            title="P-LAB Atualizar",
            warn_icon=True
        )
        return

    v_local  = versao_local.get("version", "0.0.0")
    v_remota = versao_remota.get("version", "0.0.0")

    # 2. Compara versoes
    if not versao_maior(v_remota, v_local):
        forms.alert(
            "Voce ja esta com a versao mais recente!\n\n"
            "Versao instalada: {}\n"
            "Versao disponivel: {}".format(v_local, v_remota),
            title="P-LAB Atualizar"
        )
        return

    # 3. Pergunta se quer atualizar
    notas = versao_remota.get("notas", "")
    confirmar = forms.alert(
        "Nova versao disponivel!\n\n"
        "Instalada:   {}\n"
        "Disponivel:  {}\n\n"
        "{}\n\n"
        "Deseja atualizar agora?".format(v_local, v_remota, notas),
        title="P-LAB Atualizar",
        yes=True,
        no=True
    )
    if not confirmar:
        return

    # 4. Baixa e instala
    pasta_temp = tempfile.mkdtemp()
    try:
        pasta_nova = baixar_e_extrair_zip(pasta_temp)

        if not pasta_nova:
            forms.alert(
                "Falha ao baixar a atualizacao.\n"
                "Verifique sua conexao e tente novamente.",
                title="P-LAB Atualizar",
                warn_icon=True
            )
            return

        sucesso = copiar_atualizacao(pasta_nova)

        if sucesso:
            forms.alert(
                "Atualizacao concluida!\n\n"
                "Versao instalada: {}\n\n"
                "Clique em 'Reload' na aba pyRevit para\n"
                "aplicar as mudancas.".format(v_remota),
                title="P-LAB Atualizar"
            )
        else:
            forms.alert(
                "Ocorreu um erro durante a atualizacao.\n"
                "Entre em contato com o suporte P-LAB:\n"
                "(61) 98206-8746",
                title="P-LAB Atualizar",
                warn_icon=True
            )
    finally:
        # Limpa arquivos temporarios
        try:
            shutil.rmtree(pasta_temp)
        except Exception:
            pass


main()
