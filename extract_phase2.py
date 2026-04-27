import os
import ast

def get_ast_map(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        source = f.read()
    
    tree = ast.parse(source)
    methods = []
    class_start = None
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == 'RoboApp':
            class_start = node.lineno
            for item in node.body:
                if isinstance(item, ast.FunctionDef):
                    methods.append({
                        'name': item.name,
                        'start': item.lineno,
                        'end': getattr(item, "end_lineno", item.lineno)
                    })
    return methods, source.split('\n'), class_start

def extract_mixins():
    filepath = 'robo_comprovantes_v14_interface.py'
    methods, lines, class_start = get_ast_map(filepath)
    
    os.makedirs('gui/mixins', exist_ok=True)
    with open('gui/__init__.py', 'w', encoding='utf-8') as f: f.write('')
    with open('gui/mixins/__init__.py', 'w', encoding='utf-8') as f: f.write('')
    
    mixins = {
        'gui/mixins/aba_robo.py': ('AbaRoboMixin', ['_build_aba_robo', '_iniciar_robo', '_rodar_em_thread', '_finalizar', '_abrir_saida']),
        'gui/mixins/aba_busca.py': ('AbaBuscaMixin', ['_build_aba_busca', '_executar_busca', '_busca_popular_tree', '_busca_ordenar', '_busca_abrir_pdf', '_baixar_selecionados', '_reclassificar_comprovantes', '_enviar_selecionados', '_abrir_dialogo_periodo', '_reenviar_relatorio_periodo_datas', '_reenviar_relatorio_periodo']),
        'gui/mixins/aba_extrato.py': ('AbaExtratoMixin', ['_build_aba_extrato', '_selecionar_extrato', '_processar_extrato', '_processar_fluxo_consolidado', '_enviar_dashboard_email']),
        'gui/mixins/aba_contas_pagar.py': ('AbaContasPagarMixin', ['_build_aba_contas_pagar', '_cap_consultar_sefaz', '_cap_importar_xml', '_cap_importar_xml_path', '_cap_build_lista', '_cap_carregar_lista', '_cap_mudar_status', '_cap_build_nova_nota', '_cap_calcular_parcelas', '_cap_salvar_nota', '_cap_build_adiantamentos', '_cap_build_prestacao', '_cap_build_impostos', '_imp_carregar', '_imp_marcar_pago', '_cap_build_fornecedores', '_forn_carregar', '_forn_novo', '_forn_editar', '_forn_inativar', '_forn_excluir', '_forn_abrir_form', '_cap_build_cnab', '_cnab_selecionar_pasta', '_cnab_carregar_preview', '_cnab_gerar']),
        'gui/mixins/aba_recebiveis.py': ('AbaRecebiveisMixin', ['_build_aba_recebiveis', '_card_credencial', '_resultado_box', '_log_resultado', '_limpar_resultado', '_checar_requests', '_build_sub_pagarme', '_consultar_pagarme', '_saldo_pagarme', '_build_sub_rede', '_rede_obter_token', '_consultar_rede_vendas', '_consultar_rede_recebiveis', '_build_sub_getnet', '_getnet_obter_token', '_consultar_getnet_vendas', '_consultar_getnet_recebiveis', '_build_sub_microvix', '_microvix_xml_request', '_consultar_microvix_vendas', '_consultar_microvix_estoque']),
        'gui/mixins/aba_autorizacao.py': ('AbaAutorizacaoMixin', ['_build_aba_autorizacao', '_aut_remover_linha', '_aut_ordenar', '_aut_selecionar_arquivo', '_aut_sugerir_periodo', '_aut_carregar', '_aut_popular_treeview', '_aut_editar_linha', '_aut_aplicar_edicao', '_gerar_pdf_transferencias', '_gerar_pdf_transferencia_filial', '_aut_gerar_pdf', '_gerar_pdf_transferencias_filiais', '_enviar_para_autorizacao']),
        'gui/mixins/aba_historico_status.py': ('AbaHistStatusMixin', ['_build_aba_historico_status', '_carregar_lista_logs', '_carregar_log_selecionado', '_apagar_log_selecionado', '_apagar_todos_logs']),
        'gui/mixins/aba_historico_fluxo.py': ('AbaHistFluxoMixin', ['_build_aba_historico_fluxo', '_carregar_historico_fluxo', '_filtrar_fluxo', '_ordenar_fluxo', '_atualizar_cards_fluxo', '_exportar_fluxo_excel']),
        'gui/mixins/utils_mixin.py': ('AppUtilsMixin', ['_stat_card', '_contar_pdfs', '_log', '_atualizar_stats'])
    }
    
    common_imports = """import os
import re
import sys
import shutil
import time
import threading
import sqlite3
import json
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import datetime, date
import zipfile
import tempfile
import smtplib

from config import *
from utils.data_processing import *
from utils.ocr_utils import *
from utils.email_service import *
from utils.file_manager import *
from utils.integrations_utils import *
from utils.extratos_processor import *
from utils.dashboard_service import *
from utils.formatters import *
from database import *
from cnab_generator import *

"""
    
    extracted_methods = set()
    written_files = {}
    
    for mod_path, (class_name, func_names) in mixins.items():
        mod_lines = [common_imports, f"class {class_name}:"]
        
        for fname in func_names:
            m = next((m for m in methods if m['name'] == fname), None)
            if m:
                extracted_methods.add(fname)
                idx = methods.index(m)
                
                start_line = m['start']
                if idx > 0:
                    prev_node = methods[idx-1]
                    start_limit = prev_node['end'] + 1
                else:
                    start_limit = class_start + 1
                
                actual_start = start_line
                while actual_start > start_limit:
                    line_content = lines[actual_start-2].strip()
                    if line_content.startswith('#') or line_content.startswith('@') or not line_content:
                        actual_start -= 1
                    else:
                        break
                        
                func_lines = lines[actual_start-1:m['end']]
                mod_lines.extend(func_lines)
                mod_lines.append('')
        
        written_files[mod_path] = mod_lines

    remaining_methods = [m for m in methods if m['name'] not in extracted_methods]
    
    app_lines = [common_imports]
    app_lines.append("from gui.mixins.aba_robo import AbaRoboMixin")
    app_lines.append("from gui.mixins.aba_busca import AbaBuscaMixin")
    app_lines.append("from gui.mixins.aba_extrato import AbaExtratoMixin")
    app_lines.append("from gui.mixins.aba_contas_pagar import AbaContasPagarMixin")
    app_lines.append("from gui.mixins.aba_recebiveis import AbaRecebiveisMixin")
    app_lines.append("from gui.mixins.aba_autorizacao import AbaAutorizacaoMixin")
    app_lines.append("from gui.mixins.aba_historico_status import AbaHistStatusMixin")
    app_lines.append("from gui.mixins.aba_historico_fluxo import AbaHistFluxoMixin")
    app_lines.append("from gui.mixins.utils_mixin import AppUtilsMixin")
    app_lines.append("")
    
    app_lines.append("class RoboApp(")
    app_lines.append("    AbaRoboMixin,")
    app_lines.append("    AbaBuscaMixin,")
    app_lines.append("    AbaExtratoMixin,")
    app_lines.append("    AbaContasPagarMixin,")
    app_lines.append("    AbaRecebiveisMixin,")
    app_lines.append("    AbaAutorizacaoMixin,")
    app_lines.append("    AbaHistStatusMixin,")
    app_lines.append("    AbaHistFluxoMixin,")
    app_lines.append("    AppUtilsMixin")
    app_lines.append("):")
    
    for m in remaining_methods:
        idx = methods.index(m)
        start_line = m['start']
        if idx > 0:
            prev_node = methods[idx-1]
            start_limit = prev_node['end'] + 1
        else:
            start_limit = class_start + 1
        
        actual_start = start_line
        while actual_start > start_limit:
            line_content = lines[actual_start-2].strip()
            if line_content.startswith('#') or line_content.startswith('@') or not line_content:
                actual_start -= 1
            else:
                break
                
        func_lines = lines[actual_start-1:m['end']]
        app_lines.extend(func_lines)
        app_lines.append('')
        
    for mod_path, mod_lines in written_files.items():
        with open(mod_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(mod_lines))
            
    with open('gui/app_main.py', 'w', encoding='utf-8') as f:
        f.write('\n'.join(app_lines))

    main_py_lines = [
        "import tkinter as tk",
        "from gui.app_main import RoboApp",
        "",
        "if __name__ == '__main__':",
        "    root = tk.Tk()",
        "    app = RoboApp(root)",
        "    root.mainloop()"
    ]
    with open('main.py', 'w', encoding='utf-8') as f:
        f.write('\n'.join(main_py_lines))

if __name__ == "__main__":
    extract_mixins()
