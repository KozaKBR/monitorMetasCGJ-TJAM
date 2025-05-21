# monitor_metas_corregedorias_GUI_v2.9_meta1_formula_fix.py
# Autor: Cristhiano Leite dos Santos
# Data: 20/05/2025
# Versão: 2.9 (Correção Fórmula Percentual Meta 1)
# Descrição: Calcula Metas Nacionais 1, 2 e 3 das Corregedorias 2025 (Global e por Juiz Auxiliar)
#            e gera relatório Excel com Sumário, Abas por Indicador Px.y
#            e Abas de Ação com a Tarefa Atual dos processos pendentes.

import pandas as pd
import os
from datetime import datetime, timedelta
import logging
import io
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import requests 

# --- Configuração do Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ==================================
# Classe ConfiguracaoMetas
# (Sem alterações nesta classe em relação à v2.8)
# ==================================
class ConfiguracaoMetas:
    ANO_META = 2025
    DATA_INICIO_META = datetime(ANO_META, 1, 1); DATA_FIM_META = datetime(ANO_META, 12, 31)
    DATA_CORTE_META2 = datetime(ANO_META - 1, 8, 31); DATA_FIM_ANO_ANTERIOR = datetime(ANO_META - 1, 12, 31)
    PRAZO_DIAS_META3 = 140

    MOV_ARQUIVAMENTO_DEFINITIVO = '246'; MOV_DESARQUIVAMENTO = '893'
    MOV_DETERMINACAO_ARQUIVAMENTO_1 = '1063'; MOV_DETERMINACAO_ARQUIVAMENTO_2 = '12430'
    MOV_PEDIDO_PAUTA_VOTO = '12311'

    _MOVIMENTOS_JULGAMENTO_PASTA_193 = {
        '193', '218', '228', '472', '473', '230', '235', '236', '456', '454', '457', '458', '459', '460', '461', '462', '463', '464', '11374', '11375', '11376', '11377', '11378', '11379', '11380', '11381', '12256', '12298', '12325', '14848', '15245', '15249', '15250', '15251', '853', '10953', '10961', '11373', '12319', '12458', '12459', '12709', '12710', '12711', '12712', '12713', '12714', '12715', '12716', '12717', '12718', '12719', '12720', '12721', '12722', '12723', '12724', '14218', '15253', '15254', '15255', '15256', '15257', '15258', '15259', '15260', '15261', '15262', '15263', '15264', '15265', '15266', '15408', '385', '196', '198', '200', '202', '208', '210', '442', '443', '444', '445', '12032', '12041', '12475', '212', '446', '447', '448', '449', '214', '450', '451', '452', '453', '14680', '219', '220', '221', '237', '238', '239', '240', '241', '242', '455', '466', '471', '871', '901', '972', '973', '1042', '1043', '1044', '1046', '1047', '1048', '1049', '1050', '11411', '11801', '11878', '11879', '12028', '12616', '12735', '15322', '10964', '11401', '11402', '11403', '11404', '11405', '11406', '11407', '11408', '11409', '11795', '11796', '11876', '11877', '12033', '12034', '12187', '12252', '12253', '12254', '12257', '12258', '12321', '12322', '12323', '12324', '12326', '12327', '12328', '12329', '12330', '12331', '12433', '12434', '12435', '12436', '12438', '12439', '12440', '12441', '12442', '12443', '12450', '12451', '12452', '12453', '12649', '12650', '12651', '12652', '12653', '12654', '12661', '12666', '12667', '12668', '12669', '12670', '12672', '12673', '12674', '12675', '12676', '12677', '12792', '12664', '12678', '12660', '12662', '12663', '12679', '12680', '12681', '12682', '12683', '12684', '12685', '12686', '12687', '12688', '12689', '12690', '12691', '12692', '12693', '12694', '12695', '12696', '12697', '12698', '12699', '12700', '12701', '12702', '12703', '14210', '14211', '14213', '14214', '14215', '14216', '14217', '15023', '15024', '12738', '14099', '14219', '14777', '14778', '14937', '15022', '15026', '15027', '15028', '15029', '15030', '15165', '15166', '15211', '15212', '15213', '15214', '15252'
    }
    MOVIMENTOS_DECISAO = {MOV_DETERMINACAO_ARQUIVAMENTO_1, MOV_DETERMINACAO_ARQUIVAMENTO_2, MOV_PEDIDO_PAUTA_VOTO}.union(_MOVIMENTOS_JULGAMENTO_PASTA_193)
    MOVIMENTOS_BAIXA = {MOV_ARQUIVAMENTO_DEFINITIVO}
    MOVIMENTOS_TERMINAIS = MOVIMENTOS_DECISAO.union(MOVIMENTOS_BAIXA)

    _ASSUNTOS_AGR_REC = [
        '11336', '10894', '10225', '11952', '12589', '10010', '10187', '30000009', '30000010', '10012', '11560', '11937', '30000011', '30000012', '10013', '30000013', '10011', '30000020', '30000014', '30000015', '11951', '30000024', '11950', '10881', '11915', '11916', '30000016', '10283', '30000017', '30000018', '30000019', '10014', '11919', '15072', '10949'
    ]
    CLASSES_ASSUNTOS_RELEVANTES = {
        '200': _ASSUNTOS_AGR_REC, '1299': _ASSUNTOS_AGR_REC,
        '1262': ['__TODOS__'], '1264': ['__TODOS__'], '20000002': ['__TODOS__'],
        '1301': ['__TODOS__'], '1308': ['__TODOS__'], '11892': ['__TODOS__']
    }
    CLASSES_RELEVANTES = list(CLASSES_ASSUNTOS_RELEVANTES.keys())
    CLASSE_EXCLUIDA_NOME = "Representação por Excesso de Prazo"; CLASSE_EXCLUIDA_CODIGO = None

    COLUNA_ID_PROCESSO = 'id_processo_trf'; COLUNA_NR_PROCESSO = 'nr_processo'
    COLUNA_CLASSE_COD = 'cd_classe_judicial'; COLUNA_CLASSE_NOME = 'ds_classe_judicial'
    COLUNA_ASSUNTO_COD = 'cd_assunto_principal'; COLUNA_ASSUNTO_NOME = 'ds_assunto_principal'
    COLUNA_DATA_AUTUACAO = 'dt_autuacao'
    COLUNA_MOVIMENTO_COD = 'codigo_movimento'; COLUNA_MOVIMENTO_NOME = 'movimento'
    COLUNA_MOVIMENTO_DATA = 'lancado_em'
    COLUNA_TAREFA_ID_PROCESSO = 'id_processo_trf'; COLUNA_TAREFA_FLUXO = 'fluxo'
    COLUNA_TAREFA_NOME = 'tarefa'; COLUNA_TAREFA_INICIO = 'inicio_tarefa'; COLUNA_TAREFA_FIM = 'fim_tarefa'
    
    COLUNA_JUIZ_AUXILIAR = 'Juiz Auxiliar Designado'
    API_URL_JUIZES_AUXILIARES = "https://distribuicao.tjam.jus.br/api/distribuicoes/get-all"
    DEFAULT_JUIZ_NAO_DESIGNADO = "Não Designado / API"


    COLUNAS_ESSENCIAIS_CABECALHO = [COLUNA_ID_PROCESSO, COLUNA_NR_PROCESSO, COLUNA_CLASSE_COD, COLUNA_CLASSE_NOME, COLUNA_ASSUNTO_COD, COLUNA_DATA_AUTUACAO]
    COLUNAS_ESSENCIAIS_MOVIMENTOS = [COLUNA_ID_PROCESSO, COLUNA_MOVIMENTO_COD, COLUNA_MOVIMENTO_DATA]
    COLUNAS_ESSENCIAIS_TAREFAS = [COLUNA_TAREFA_ID_PROCESSO, COLUNA_TAREFA_FLUXO, COLUNA_TAREFA_NOME, COLUNA_TAREFA_INICIO, COLUNA_TAREFA_FIM]

# ==========================
# Classe CarregadorDados
# (Sem alterações nesta classe em relação à v2.8)
# ==========================
class CarregadorDados:
    def __init__(self, config=ConfiguracaoMetas, logger=None):
        self.config = config
        self.logger = logger if logger else logging.getLogger()

    def _validar_colunas(self, df, colunas_essenciais, nome_arquivo):
        if df is None: return False
        df_cols_lower = [str(col).lower() for col in df.columns]
        essenciais_lower = [str(col).lower() for col in colunas_essenciais]
        colunas_faltantes = [col for col in essenciais_lower if col not in df_cols_lower]
        if colunas_faltantes:
            self.logger.error(f"Colunas essenciais não encontradas em '{os.path.basename(nome_arquivo)}': {', '.join(colunas_faltantes)}")
            self.logger.error(f"      Colunas disponíveis ({len(df.columns)}): {', '.join(map(str, df.columns))}")
            return False
        return True

    def _renomear_colunas_para_padrao(self, df, colunas_essenciais):
        mapa_renomear = {}
        df_cols_lower_map = {str(col).lower(): str(col) for col in df.columns}
        for col_padrao in colunas_essenciais:
            col_padrao_lower = str(col_padrao).lower();
            if col_padrao_lower in df_cols_lower_map:
                col_original = df_cols_lower_map[col_padrao_lower];
                if col_original != col_padrao: mapa_renomear[col_original] = col_padrao
        if mapa_renomear: self.logger.info(f"Renomeando colunas: {mapa_renomear}"); df.rename(columns=mapa_renomear, inplace=True)
        return df

    def carregar_arquivo(self, caminho_arquivo, colunas_essenciais, tipo_arquivo="Dados"):
        if not caminho_arquivo or not os.path.exists(caminho_arquivo): self.logger.error(f"Arquivo {tipo_arquivo} não encontrado: {caminho_arquivo}"); return None
        nome_arquivo = os.path.basename(caminho_arquivo); df = None; encodings = ['latin-1', 'windows-1252', 'utf-8', 'utf-8-sig']
        try:
            if nome_arquivo.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(caminho_arquivo, engine=None, dtype=str); self.logger.info(f"'{nome_arquivo}' ({tipo_arquivo}) Excel carregado.")
            elif nome_arquivo.lower().endswith('.csv'):
                self.logger.info(f"Carregando CSV {tipo_arquivo}: {nome_arquivo}...")
                for enc in encodings:
                    try:
                        with open(caminho_arquivo, 'r', encoding=enc) as f: preview=f.readline()+f.readline()
                        sep = ';' if preview.count(';') >= preview.count(',') else ','
                        self.logger.debug(f"  Tentando CSV: enc='{enc}', sep='{repr(sep)}'")
                        df = pd.read_csv(caminho_arquivo, encoding=enc, sep=sep, on_bad_lines='warn', low_memory=False, dtype=str)
                        self.logger.info(f"'{nome_arquivo}' ({tipo_arquivo}) CSV carregado (enc={enc}, sep='{repr(sep)}')."); break
                    except Exception: self.logger.debug(f"  Falha com enc='{enc}'."); continue
                if df is None: self.logger.error(f"ERRO: Falha ao carregar CSV '{nome_arquivo}'."); return None
            else: self.logger.error(f"ERRO: Formato não suportado p/ {tipo_arquivo}: '{nome_arquivo}'."); return None

            if not self._validar_colunas(df, colunas_essenciais, nome_arquivo): return None
            df = self._renomear_colunas_para_padrao(df, colunas_essenciais)
            self.logger.info(f"'{nome_arquivo}' ({tipo_arquivo}): {df.shape[0]} linhas.")
            self.logger.info(f"Convertendo tipos para {tipo_arquivo}...")

            cols_to_numeric = []; cols_to_datetime = []; cols_to_string_special = []
            if tipo_arquivo == "Cabeçalhos":
                cols_to_numeric = [self.config.COLUNA_ID_PROCESSO]
                cols_to_datetime = [self.config.COLUNA_DATA_AUTUACAO]
                cols_to_string_special = [self.config.COLUNA_NR_PROCESSO, self.config.COLUNA_CLASSE_COD, self.config.COLUNA_ASSUNTO_COD, self.config.COLUNA_CLASSE_NOME]
            elif tipo_arquivo == "Movimentos":
                cols_to_numeric = [self.config.COLUNA_ID_PROCESSO]
                cols_to_datetime = [self.config.COLUNA_MOVIMENTO_DATA]
                cols_to_string_special = [self.config.COLUNA_MOVIMENTO_COD]
            elif tipo_arquivo == "Tarefas":
                cols_to_numeric = [self.config.COLUNA_TAREFA_ID_PROCESSO]
                cols_to_datetime = [self.config.COLUNA_TAREFA_INICIO, self.config.COLUNA_TAREFA_FIM]
                cols_to_string_special = [self.config.COLUNA_TAREFA_FLUXO, self.config.COLUNA_TAREFA_NOME]

            for col in df.columns:
                if df[col].dtype == 'object': df[col] = df[col].astype(str).str.strip()
                if col in cols_to_numeric: df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')
                elif col in cols_to_datetime: df[col] = pd.to_datetime(df[col], dayfirst=False, errors='coerce')
                elif col in cols_to_string_special: df[col] = df[col].fillna('').astype(str)

            self.logger.info(f"Conversões {tipo_arquivo} concluídas.")
            return df
        except Exception as e:
            self.logger.error(f"ERRO GERAL carregando {tipo_arquivo} '{nome_arquivo}': {e}")
            self.logger.error(traceback.format_exc()); return None

# =======================================
# Classe IdentificadorProcessosMeta
# (Sem alterações nesta classe em relação à v2.8)
# =======================================
class IdentificadorProcessosMeta:
    def __init__(self, config=ConfiguracaoMetas, logger=None):
        self.config = config
        self.logger = logger if logger else logging.getLogger()

    def identificar(self, df_cabecalhos):
        if df_cabecalhos is None or df_cabecalhos.empty: self.logger.error("ERRO: DF Cabeçalhos vazio."); return pd.Series(dtype='Int64')
        self.logger.info("Identificando processos relevantes...")
        req_cols = [self.config.COLUNA_ID_PROCESSO, self.config.COLUNA_CLASSE_COD, self.config.COLUNA_CLASSE_NOME, self.config.COLUNA_ASSUNTO_COD]
        if not all(c in df_cabecalhos.columns for c in req_cols): self.logger.error(f"ERRO: Faltam colunas p/ ID ({', '.join(req_cols)})."); return pd.Series(dtype='Int64')
        try:
            for col in [self.config.COLUNA_CLASSE_COD, self.config.COLUNA_ASSUNTO_COD, self.config.COLUNA_CLASSE_NOME]: df_cabecalhos[col] = df_cabecalhos[col].fillna('').astype(str)
        except Exception as e: self.logger.error(f"ERRO conversão códigos/nomes: {e}"); return pd.Series(dtype='Int64')

        mascara_relevante = pd.Series([False]*len(df_cabecalhos), index=df_cabecalhos.index)
        self.logger.info("Aplicando regras de inclusão (Classe/Assunto)...")
        for classe_cod, assuntos in self.config.CLASSES_ASSUNTOS_RELEVANTES.items():
            mascara_classe = (df_cabecalhos[self.config.COLUNA_CLASSE_COD] == classe_cod)
            if not mascara_classe.any(): continue
            if assuntos == ['__TODOS__']: 
                mascara_relevante |= mascara_classe
            elif isinstance(assuntos, list) and assuntos:
                mascara_assunto = df_cabecalhos[self.config.COLUNA_ASSUNTO_COD].isin(assuntos)
                mascara_relevante |= (mascara_classe & mascara_assunto)
        self.logger.info(f"Total após inclusão: {mascara_relevante.sum()}")

        mascara_exclusao = pd.Series([False]*len(df_cabecalhos), index=df_cabecalhos.index)
        if self.config.CLASSE_EXCLUIDA_CODIGO: mascara_exclusao = (df_cabecalhos[self.config.COLUNA_CLASSE_COD] == self.config.CLASSE_EXCLUIDA_CODIGO)
        elif self.config.CLASSE_EXCLUIDA_NOME: mascara_exclusao = (df_cabecalhos[self.config.COLUNA_CLASSE_NOME].str.lower() == self.config.CLASSE_EXCLUIDA_NOME.lower())
        
        num_a_excluir = 0
        if mascara_exclusao.any():
            num_a_excluir = (mascara_relevante & mascara_exclusao).sum()
            self.logger.info(f"Aplicando exclusão ({num_a_excluir} processos marcados para exclusão)..")
            mascara_final = mascara_relevante & (~mascara_exclusao)
        else:
            mascara_final = mascara_relevante
        
        self.logger.info(f"Total relevantes final: {mascara_final.sum()}")


        ids_rel = df_cabecalhos.loc[mascara_final, self.config.COLUNA_ID_PROCESSO].unique()
        ids_rel_series = pd.Series(ids_rel).dropna().astype('Int64')
        return ids_rel_series

# ==========================
# Classe CalculadoraMeta1
# ==========================
class CalculadoraMeta1:
    def __init__(self, config=ConfiguracaoMetas, logger=None):
        self.config = config; self.logger = logger if logger else logging.getLogger()

    def _find_definitive_archives(self, df_mov_proc, date_limit=None):
        mov_arq=self.config.MOV_ARQUIVAMENTO_DEFINITIVO; mov_desarq=self.config.MOV_DESARQUIVAMENTO
        col_data=self.config.COLUNA_MOVIMENTO_DATA; col_cod=self.config.COLUNA_MOVIMENTO_COD
        if df_mov_proc is None or df_mov_proc.empty: return None
        df_mov_work = df_mov_proc.copy()
        df_mov_work[col_data] = pd.to_datetime(df_mov_work[col_data], errors='coerce')
        if date_limit and isinstance(date_limit, datetime): df_mov_work = df_mov_work.loc[df_mov_work[col_data] <= date_limit]
        if df_mov_work.empty: return None
        last_arq = df_mov_work.loc[df_mov_work[col_cod] == mov_arq]; last_desarq = df_mov_work.loc[df_mov_work[col_cod] == mov_desarq]
        if last_arq.empty: return None
        max_dt_arq = last_arq[col_data].max()
        if pd.isna(max_dt_arq): return None
        if last_desarq.empty: return max_dt_arq
        max_dt_desarq = last_desarq[col_data].max()
        if pd.isna(max_dt_desarq) or max_dt_arq > max_dt_desarq: return max_dt_arq
        else: return None

    def calcular(self, df_cabecalhos_meta_especifica, df_movimentos_meta_especifica, ids_relevantes_meta_especifica):
        df_cabecalhos = df_cabecalhos_meta_especifica
        df_movimentos = df_movimentos_meta_especifica
        ids_relevantes = ids_relevantes_meta_especifica
        
        default_result = {'P1.1': 0, 'P1.2': 0, 'P1.3': 0, 'percentual': 0.0, 'ids_P1.1': [], 'ids_P1.2': [], 'ids_P1.3': []}
        if df_cabecalhos is None or df_movimentos is None or ids_relevantes is None or not hasattr(ids_relevantes, 'empty'): 
            self.logger.debug("Meta1: Inputs inválidos ou ids_relevantes não é uma Series, retornando default."); # Mudado para debug
            return default_result
        
        self.logger.debug("Calculando Meta 1..."); col_id = self.config.COLUNA_ID_PROCESSO
        
        if ids_relevantes.empty: 
            self.logger.debug("Meta1: ids_relevantes vazio, retornando default.");
            return default_result

        # P1.1: Distribuídos no ano da meta (2025)
        data_inicio_meta = self.config.DATA_INICIO_META; col_data_aut = self.config.COLUNA_DATA_AUTUACAO
        mask_p1_1 = df_cabecalhos[col_data_aut] >= data_inicio_meta
        ids_p1_1 = df_cabecalhos.loc[mask_p1_1, col_id].unique().tolist(); p1_1 = len(ids_p1_1); self.logger.debug(f"P1.1 = {p1_1}")

        # P1.3: Saldo do ano anterior (processos relevantes pendentes ao final de 2024)
        data_fim_ano_anterior = self.config.DATA_FIM_ANO_ANTERIOR; ids_arq_def_fim_2024 = set(); mov_grouped = df_movimentos.groupby(col_id)
        for proc_id in ids_relevantes: 
            df_proc_mov = mov_grouped.get_group(proc_id) if proc_id in mov_grouped.groups else pd.DataFrame()
            dt_arq = self._find_definitive_archives(df_proc_mov.copy(), data_fim_ano_anterior)
            if dt_arq is not None: ids_arq_def_fim_2024.add(proc_id)
        
        ids_p1_3 = list(set(ids_relevantes.tolist()) - ids_arq_def_fim_2024); p1_3 = len(ids_p1_3); self.logger.debug(f"P1.3 = {p1_3}")

        # P1.2: Baixados/Resolvidos no ano da meta (2025) - inclui de P1.1 e P1.3
        ids_p1_2 = []
        for proc_id in ids_relevantes: 
            df_proc_mov = mov_grouped.get_group(proc_id) if proc_id in mov_grouped.groups else pd.DataFrame()
            dt_arq_def_final = self._find_definitive_archives(df_proc_mov.copy()) 
            if dt_arq_def_final is not None and dt_arq_def_final >= data_inicio_meta and dt_arq_def_final <= self.config.DATA_FIM_META: # Assegura que é dentro do ano da meta
                 ids_p1_2.append(proc_id)
        p1_2 = len(ids_p1_2); self.logger.debug(f"P1.2 = {p1_2} (Total baixados def. em {self.config.ANO_META})")

        # <<< CORREÇÃO: Nova fórmula para o percentual da Meta 1 >>>
        percentual = 0.0
        denominador_novo = p1_1 + 1 
        # Denominador sempre será >= 1 porque p1_1 (contagem) >= 0.
        percentual = (p1_2 / denominador_novo) * 100
        
        self.logger.debug(f"Percentual Meta 1 (fórmula P1.2/(P1.1+1)) = {percentual:.2f}%")
        return {'P1.1': p1_1, 'P1.2': p1_2, 'P1.3': p1_3, 'percentual': round(percentual, 2), 'ids_P1.1': ids_p1_1, 'ids_P1.2': ids_p1_2, 'ids_P1.3': ids_p1_3}

# ==========================
# Classe CalculadoraMeta2
# (Sem alterações nesta classe em relação à v2.8)
# ==========================
class CalculadoraMeta2:
    def __init__(self, config=ConfiguracaoMetas, logger=None):
        self.config = config; self.logger = logger if logger else logging.getLogger()

    def _get_terminal_status(self, df_mov_proc, date_limit):
        if df_mov_proc is None or df_mov_proc.empty: return 'PENDING'
        mov_arq=self.config.MOV_ARQUIVAMENTO_DEFINITIVO; mov_desarq=self.config.MOV_DESARQUIVAMENTO
        mov_terminais=self.config.MOVIMENTOS_TERMINAIS; col_data=self.config.COLUNA_MOVIMENTO_DATA; col_cod=self.config.COLUNA_MOVIMENTO_COD
        df_mov_work = df_mov_proc.copy(); df_mov_work[col_data] = pd.to_datetime(df_mov_work[col_data], errors='coerce')
        df_mov_limitado = df_mov_work.loc[df_mov_work[col_data] <= date_limit].copy()
        if df_mov_limitado.empty: return 'PENDING'
        df_terminais_limitado = df_mov_limitado[df_mov_limitado[col_cod].isin(mov_terminais)].copy()
        if df_terminais_limitado.empty: return 'PENDING'
        try:
            idx_ultimo_terminal = df_terminais_limitado[col_data].idxmax()
            if pd.isna(idx_ultimo_terminal): return 'PENDING' 
            ultimo_terminal_event = df_terminais_limitado.loc[idx_ultimo_terminal]
        except ValueError: return 'PENDING' 
        
        ultimo_terminal_code = ultimo_terminal_event[col_cod]; ultimo_terminal_date = ultimo_terminal_event[col_data]
        if pd.isna(ultimo_terminal_date): return 'PENDING' 
        
        if ultimo_terminal_code != mov_arq: return 'DECIDED_OTHER' 
        
        desarquivamentos_posteriores = df_mov_limitado[(df_mov_limitado[col_cod] == mov_desarq) & (df_mov_limitado[col_data] > ultimo_terminal_date)]
        return 'PENDING' if not desarquivamentos_posteriores.empty else 'ARCHIVED_DEFINITIVE'


    def calcular(self, df_cabecalhos_meta_especifica, df_movimentos_meta_especifica, ids_relevantes_meta_especifica):
        df_cabecalhos = df_cabecalhos_meta_especifica
        df_movimentos = df_movimentos_meta_especifica
        ids_relevantes = ids_relevantes_meta_especifica
        
        default_result = {'P2.1': 0, 'P2.2': 0, 'percentual': 100.0, 'ids_P2.1': [], 'ids_P2.2': []}
        if df_cabecalhos is None or df_movimentos is None or ids_relevantes is None or not hasattr(ids_relevantes, 'empty'): 
            self.logger.debug("Meta2: Inputs inválidos, retornando default.") # Mudado para debug
            return default_result
        self.logger.debug("Calculando Meta 2..."); col_id = self.config.COLUNA_ID_PROCESSO
        
        if ids_relevantes.empty: 
            self.logger.debug("Meta2: ids_relevantes vazio, retornando default.")
            return default_result

        data_corte_meta2 = self.config.DATA_CORTE_META2; col_data_aut = self.config.COLUNA_DATA_AUTUACAO
        mask_p2_1_cand = df_cabecalhos[col_data_aut] <= data_corte_meta2
        ids_candidatos_p2_1 = df_cabecalhos.loc[mask_p2_1_cand, col_id].unique()
        self.logger.debug(f"Candidatos P2.1 (autuados até {data_corte_meta2.strftime('%Y-%m-%d')}): {len(ids_candidatos_p2_1)}")
        if len(ids_candidatos_p2_1) == 0: return default_result

        ids_p2_1_list = []; data_fim_ano_anterior = self.config.DATA_FIM_ANO_ANTERIOR
        mov_cand_grouped = df_movimentos[df_movimentos[col_id].isin(ids_candidatos_p2_1)].groupby(col_id)
        for proc_id in ids_candidatos_p2_1:
            df_proc_mov = mov_cand_grouped.get_group(proc_id) if proc_id in mov_cand_grouped.groups else pd.DataFrame()
            status = self._get_terminal_status(df_proc_mov.copy(), data_fim_ano_anterior)
            if status == 'PENDING': ids_p2_1_list.append(proc_id)
        p2_1 = len(ids_p2_1_list); self.logger.debug(f"P2.1 = {p2_1}")
        if p2_1 == 0: return default_result

        ids_p2_2_list = []; data_fim_meta = self.config.DATA_FIM_META
        mov_p2_1_grouped = df_movimentos[df_movimentos[col_id].isin(ids_p2_1_list)].groupby(col_id)
        for proc_id in ids_p2_1_list:
            df_proc_mov = mov_p2_1_grouped.get_group(proc_id) if proc_id in mov_p2_1_grouped.groups else pd.DataFrame()
            status_final = self._get_terminal_status(df_proc_mov.copy(), data_fim_meta)
            if status_final != 'PENDING': ids_p2_2_list.append(proc_id)
        p2_2 = len(ids_p2_2_list); self.logger.debug(f"P2.2 = {p2_2}")

        percentual = (p2_2 / p2_1) * 100 if p2_1 > 0 else 100.0; self.logger.debug(f"Percentual Meta 2 = {percentual:.2f}%")
        return {'P2.1': p2_1, 'P2.2': p2_2, 'percentual': round(percentual, 2), 'ids_P2.1': ids_p2_1_list, 'ids_P2.2': ids_p2_2_list}

# ==========================
# Classe CalculadoraMeta3
# (Sem alterações nesta classe em relação à v2.8)
# ==========================
class CalculadoraMeta3:
    def __init__(self, config=ConfiguracaoMetas, logger=None):
        self.config = config; self.logger = logger if logger else logging.getLogger()

    def _find_first_terminal_date(self, df_mov_proc, terminal_codes_set):
        if df_mov_proc is None or df_mov_proc.empty: return pd.NaT
        col_data=self.config.COLUNA_MOVIMENTO_DATA; col_cod=self.config.COLUNA_MOVIMENTO_COD
        df_mov_work = df_mov_proc.copy()
        df_mov_work[col_data] = pd.to_datetime(df_mov_work[col_data], errors='coerce')
        terminal_mask = df_mov_work[col_cod].isin(terminal_codes_set)
        df_terminal = df_mov_work.loc[terminal_mask]
        if df_terminal.empty or df_terminal[col_data].isnull().all(): return pd.NaT
        min_date = df_terminal[col_data].min() 
        return pd.NaT if pd.isna(min_date) else min_date


    def calcular(self, df_cabecalhos_meta_especifica, df_movimentos_meta_especifica, ids_relevantes_meta_especifica):
        df_cabecalhos = df_cabecalhos_meta_especifica
        df_movimentos = df_movimentos_meta_especifica
        ids_relevantes = ids_relevantes_meta_especifica
        
        default_result = {'P3.1': 0, 'P3.2': 0, 'percentual': 100.0, 'details_P3.1': {}, 'details_P3.2': {}}
        if df_cabecalhos is None or df_movimentos is None or ids_relevantes is None or not hasattr(ids_relevantes, 'empty'):
            self.logger.debug("Meta3: Inputs inválidos, retornando default.") # Mudado para debug
            return default_result
        
        self.logger.debug("Calculando Meta 3..."); col_id=self.config.COLUNA_ID_PROCESSO; col_data_aut=self.config.COLUNA_DATA_AUTUACAO

        if ids_relevantes.empty: 
            self.logger.debug("Meta3: ids_relevantes vazio, retornando default.")
            return default_result

        map_id_data_autuacao = df_cabecalhos.set_index(col_id)[col_data_aut].dropna().to_dict()
        self.logger.debug(f"Map ID->Autuação (Meta3) criado para {len(map_id_data_autuacao)} processos do escopo.")

        processos_p3_1_details = {}
        terminal_codes = self.config.MOVIMENTOS_TERMINAIS
        data_inicio_meta = self.config.DATA_INICIO_META; data_fim_meta = self.config.DATA_FIM_META
        mov_rel_grouped = df_movimentos.groupby(col_id)
        
        for proc_id in ids_relevantes: 
            if proc_id not in map_id_data_autuacao: continue 
            
            df_proc_mov = mov_rel_grouped.get_group(proc_id) if proc_id in mov_rel_grouped.groups else pd.DataFrame()
            dt_primeiro_terminal = self._find_first_terminal_date(df_proc_mov.copy(), terminal_codes)
            
            if not pd.isna(dt_primeiro_terminal) and data_inicio_meta <= dt_primeiro_terminal <= data_fim_meta:
                processos_p3_1_details[proc_id] = dt_primeiro_terminal
        p3_1 = len(processos_p3_1_details)
        self.logger.debug(f"P3.1 = {p3_1} (Decididos em {self.config.ANO_META} com data autuação válida)")
        if p3_1 == 0: return default_result

        processos_p3_2_details = {}
        prazo_maximo_dias = self.config.PRAZO_DIAS_META3
        for proc_id, dt_decisao in processos_p3_1_details.items():
            dt_autuacao = map_id_data_autuacao.get(proc_id) 
            if pd.isna(dt_autuacao) or pd.isna(dt_decisao): continue 
            delta_dias = (dt_decisao - dt_autuacao).days
            if 0 <= delta_dias <= prazo_maximo_dias:
                processos_p3_2_details[proc_id] = {'dt_autuacao': dt_autuacao, 'dt_decisao': dt_decisao, 'delta_dias': delta_dias}
        p3_2 = len(processos_p3_2_details)
        self.logger.debug(f"P3.2 = {p3_2} (Resolvidos no prazo)")

        percentual = (p3_2 / p3_1) * 100 if p3_1 > 0 else 100.0
        self.logger.debug(f"Percentual Meta 3 = {percentual:.2f}%")
        return {'P3.1': p3_1, 'P3.2': p3_2, 'percentual': round(percentual, 2),
                'details_P3.1': processos_p3_1_details, 'details_P3.2': processos_p3_2_details}

# ==========================
# Classe AnalisadorMetas
# (Sem alterações nesta classe em relação à v2.8)
# ==========================
class AnalisadorMetas:
    def __init__(self, config=ConfiguracaoMetas, logger=None):
        self.config = config
        self.logger = logger if logger else logging.getLogger()
        self.df_cabecalhos_global = None 
        self.df_movimentos_global = None
        self.df_tarefas_global = None
        self.ids_relevantes_global = None
        self.map_nr_processo_to_juiz = {}

    def _fetch_juizes_auxiliares_data(self):
        self.logger.info(f"Buscando dados de Juízes Auxiliares da API: {self.config.API_URL_JUIZES_AUXILIARES}")
        try:
            response = requests.get(self.config.API_URL_JUIZES_AUXILIARES, timeout=15)
            response.raise_for_status()
            api_data = response.json()

            if api_data.get("status") == "success" and "data" in api_data and "distribuicoes" in api_data["data"]:
                distribuicoes = api_data["data"]["distribuicoes"]
                self.logger.info(f"API retornou {len(distribuicoes)} distribuições.")
                for item in distribuicoes:
                    nr_processo = item.get("processo")
                    magistrado_info = item.get("magistrado")
                    if nr_processo and magistrado_info and isinstance(magistrado_info, dict):
                        nome_juiz = magistrado_info.get("nome")
                        if nome_juiz:
                            self.map_nr_processo_to_juiz[str(nr_processo).strip()] = str(nome_juiz).strip()
                self.logger.info(f"Mapeamento NrProcesso -> Juiz Auxiliar criado para {len(self.map_nr_processo_to_juiz)} processos distintos da API.")
                return True
            else:
                self.logger.warning("Resposta da API de Juízes Auxiliares não contém os dados esperados.")
                return False
        except requests.exceptions.Timeout:
            self.logger.error("Timeout ao buscar dados da API de Juízes Auxiliares.")
            return False
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Erro ao buscar dados da API de Juízes Auxiliares: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Erro inesperado ao processar dados da API de Juízes Auxiliares: {e}")
            self.logger.error(traceback.format_exc())
            return False

    def _augment_cabecalhos_with_juiz_info(self):
        if self.df_cabecalhos_global is None or self.config.COLUNA_NR_PROCESSO not in self.df_cabecalhos_global.columns:
            self.logger.error("DataFrame de Cabeçalhos não carregado ou sem coluna de número do processo para augmentation.")
            if self.df_cabecalhos_global is not None: 
                self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR] = self.config.DEFAULT_JUIZ_NAO_DESIGNADO 
            return

        if not self.map_nr_processo_to_juiz:
            self.logger.warning("Mapeamento de Juízes Auxiliares está vazio. Processos não serão designados via API.")
            self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR] = self.config.DEFAULT_JUIZ_NAO_DESIGNADO
            return

        self.df_cabecalhos_global[self.config.COLUNA_NR_PROCESSO] = self.df_cabecalhos_global[self.config.COLUNA_NR_PROCESSO].astype(str).str.strip()
        
        self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR] = self.df_cabecalhos_global[self.config.COLUNA_NR_PROCESSO].map(self.map_nr_processo_to_juiz)
        self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR] = self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR].fillna(self.config.DEFAULT_JUIZ_NAO_DESIGNADO)
        
        juizes_encontrados = self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR].unique()
        self.logger.info(f"Coluna '{self.config.COLUNA_JUIZ_AUXILIAR}' adicionada/atualizada em Cabeçalhos.")
        self.logger.info(f"Juízes/Categorias encontradas nos dados: {juizes_encontrados}")


    def executar_analise(self, caminho_cabecalho, caminho_movimentos, caminho_tarefas):
        self.logger.info(">>> INICIANDO ANÁLISE COMPLETA DAS METAS (GLOBAL E POR JUIZ AUXILIAR) <<<")
        
        self.logger.info("--- Etapa 1: Carregando Dados Base ---")
        carregador = CarregadorDados(config=self.config, logger=self.logger)
        self.df_cabecalhos_global = carregador.carregar_arquivo(caminho_cabecalho, self.config.COLUNAS_ESSENCIAIS_CABECALHO, "Cabeçalhos")
        if self.df_cabecalhos_global is None: return None
        self.df_movimentos_global = carregador.carregar_arquivo(caminho_movimentos, self.config.COLUNAS_ESSENCIAIS_MOVIMENTOS, "Movimentos")
        if self.df_movimentos_global is None: return None
        self.df_tarefas_global = carregador.carregar_arquivo(caminho_tarefas, self.config.COLUNAS_ESSENCIAIS_TAREFAS, "Tarefas")
        if self.df_tarefas_global is None: 
            self.logger.warning("Arquivo de Tarefas não carregado. Relatório não incluirá informações de tarefa atual.")
            self.df_tarefas_global = pd.DataFrame(columns=self.config.COLUNAS_ESSENCIAIS_TAREFAS)


        self.logger.info("--- Etapa 2: Carregando e Integrando Dados da API de Juízes Auxiliares ---")
        if not self._fetch_juizes_auxiliares_data():
            self.logger.warning("Não foi possível buscar dados da API de Juízes. Análise prosseguirá com todos como 'Não Designado'.")
        self._augment_cabecalhos_with_juiz_info() 

        self.logger.info("--- Etapa 3: Identificando Processos Relevantes (Global) ---")
        identificador = IdentificadorProcessosMeta(config=self.config, logger=self.logger)
        self.ids_relevantes_global = identificador.identificar(self.df_cabecalhos_global.copy()) 
        if self.ids_relevantes_global is None: self.logger.error("Falha na identificação global de IDs relevantes."); return None
        if self.ids_relevantes_global.empty: self.logger.warning("Nenhum processo relevante identificado globalmente.")
        
        self.logger.info(f"--- Identificados {len(self.ids_relevantes_global)} processos relevantes globalmente ---")

        self.logger.info("--- Etapa 4: Calculando Metas ---")
        todos_os_resultados = {} 

        escopos_de_calculo = ['Global']
        if self.config.COLUNA_JUIZ_AUXILIAR in self.df_cabecalhos_global.columns:
             escopos_de_calculo += sorted(list(self.df_cabecalhos_global[self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR] != self.config.DEFAULT_JUIZ_NAO_DESIGNADO][self.config.COLUNA_JUIZ_AUXILIAR].unique()))
             if self.config.DEFAULT_JUIZ_NAO_DESIGNADO in self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR].unique():
                 escopos_de_calculo.append(self.config.DEFAULT_JUIZ_NAO_DESIGNADO)
        else: 
            self.logger.warning(f"Coluna {self.config.COLUNA_JUIZ_AUXILIAR} não encontrada em df_cabecalhos_global. Apenas escopo 'Global' será calculado.")

        escopos_de_calculo = [esc for esc in list(dict.fromkeys(escopos_de_calculo)) if esc] 

        self.logger.info(f"Escopos de cálculo definidos: {escopos_de_calculo}")

        for escopo_atual in escopos_de_calculo:
            self.logger.info(f"--- Calculando Metas para o Escopo: {escopo_atual} ---")
            todos_os_resultados[escopo_atual] = {}

            ids_relevantes_escopo_atual = pd.Series(dtype='Int64')
            df_cabecalhos_escopo_atual = pd.DataFrame()

            if escopo_atual == 'Global':
                ids_relevantes_escopo_atual = self.ids_relevantes_global.copy()
                df_cabecalhos_escopo_atual = self.df_cabecalhos_global[self.df_cabecalhos_global[self.config.COLUNA_ID_PROCESSO].isin(ids_relevantes_escopo_atual)].copy()
            elif self.config.COLUNA_JUIZ_AUXILIAR in self.df_cabecalhos_global.columns: 
                mask_juiz = self.df_cabecalhos_global[self.config.COLUNA_JUIZ_AUXILIAR] == escopo_atual
                ids_escopo_temp = self.df_cabecalhos_global.loc[mask_juiz, self.config.COLUNA_ID_PROCESSO]
                ids_relevantes_escopo_atual = pd.Series(list(set(self.ids_relevantes_global) & set(ids_escopo_temp))).astype('Int64')
                df_cabecalhos_escopo_atual = self.df_cabecalhos_global[self.df_cabecalhos_global[self.config.COLUNA_ID_PROCESSO].isin(ids_relevantes_escopo_atual)].copy()
            else: 
                 self.logger.warning(f"Pulando escopo '{escopo_atual}' pois a coluna de designação de juiz não está disponível.")
                 continue


            if ids_relevantes_escopo_atual.empty:
                self.logger.warning(f"Nenhum processo relevante para o escopo: {escopo_atual}. Metas zeradas para este escopo.")
                todos_os_resultados[escopo_atual]['meta1'] = CalculadoraMeta1(config=self.config, logger=self.logger).calcular(pd.DataFrame(columns=self.df_cabecalhos_global.columns if self.df_cabecalhos_global is not None else []), pd.DataFrame(columns=self.df_movimentos_global.columns if self.df_movimentos_global is not None else []), pd.Series(dtype='Int64')) 
                todos_os_resultados[escopo_atual]['meta2'] = CalculadoraMeta2(config=self.config, logger=self.logger).calcular(pd.DataFrame(columns=self.df_cabecalhos_global.columns if self.df_cabecalhos_global is not None else []), pd.DataFrame(columns=self.df_movimentos_global.columns if self.df_movimentos_global is not None else []), pd.Series(dtype='Int64'))
                todos_os_resultados[escopo_atual]['meta3'] = CalculadoraMeta3(config=self.config, logger=self.logger).calcular(pd.DataFrame(columns=self.df_cabecalhos_global.columns if self.df_cabecalhos_global is not None else []), pd.DataFrame(columns=self.df_movimentos_global.columns if self.df_movimentos_global is not None else []), pd.Series(dtype='Int64'))
                continue

            self.logger.info(f"Escopo '{escopo_atual}': {len(ids_relevantes_escopo_atual)} IDs relevantes para cálculo.")
            
            df_movimentos_escopo_atual = self.df_movimentos_global[self.df_movimentos_global[self.config.COLUNA_ID_PROCESSO].isin(ids_relevantes_escopo_atual)].copy()
            
            calculadoras = {'meta1': CalculadoraMeta1, 'meta2': CalculadoraMeta2, 'meta3': CalculadoraMeta3}
            for nome_meta, ClasseCalc in calculadoras.items():
                try:
                    self.logger.info(f"Calculando {nome_meta.upper()} para escopo '{escopo_atual}'...")
                    calc = ClasseCalc(config=self.config, logger=self.logger)
                    res_meta = calc.calcular(df_cabecalhos_escopo_atual, 
                                             df_movimentos_escopo_atual, 
                                             ids_relevantes_escopo_atual.copy()) 
                    todos_os_resultados[escopo_atual][nome_meta] = res_meta if res_meta else {} 
                    if res_meta and res_meta.get('percentual') is not None: 
                        self.logger.info(f"{nome_meta.upper()} ({escopo_atual}) calculada: {res_meta.get('percentual', 0.0):.2f}%")
                    else: 
                        self.logger.error(f"Falha no cálculo de {nome_meta.upper()} para {escopo_atual} ou resultado inválido.")
                        if nome_meta not in todos_os_resultados[escopo_atual] or not todos_os_resultados[escopo_atual][nome_meta]: 
                             todos_os_resultados[escopo_atual][nome_meta] = {} 
                except Exception as e:
                    self.logger.error(f"ERRO INESPERADO {nome_meta.upper()} ({escopo_atual}): {e}"); self.logger.error(traceback.format_exc()); 
                    todos_os_resultados[escopo_atual][nome_meta] = {} 
        
        self.logger.info(">>> ANÁLISE COMPLETA CONCLUÍDA <<<")
        return {
            'resultados_por_escopo': todos_os_resultados, 
            'df_cabecalhos_global': self.df_cabecalhos_global, 
            'df_movimentos_global': self.df_movimentos_global,
            'df_tarefas_global': self.df_tarefas_global,
            'ids_relevantes_global': self.ids_relevantes_global,
            'escopos_calculados': escopos_de_calculo 
        }

# ==========================
# Classe GeradorRelatorio
# (Alteração no nome da aba do Excel, o resto igual à v2.8)
# ==========================
class GeradorRelatorio:
    def __init__(self, config=ConfiguracaoMetas, logger=None):
        self.config = config
        self.logger = logger if logger else logging.getLogger()
        self._calculadora_m2_helper = CalculadoraMeta2(config=config, logger=logger) 

    def _criar_map_id_nrprocesso_e_juiz(self, df_cabecalhos_global):
        col_id = self.config.COLUNA_ID_PROCESSO
        col_nr = self.config.COLUNA_NR_PROCESSO
        col_juiz = self.config.COLUNA_JUIZ_AUXILIAR 

        map_id_info = {}
        if df_cabecalhos_global is None or df_cabecalhos_global.empty:
            return map_id_info

        required_cols = [col_id, col_nr]
        if col_juiz not in df_cabecalhos_global.columns: 
            df_cabecalhos_global[col_juiz] = self.config.DEFAULT_JUIZ_NAO_DESIGNADO
        
        if not all(c in df_cabecalhos_global.columns for c in required_cols):
            self.logger.error("Faltam colunas ID ou NR_PROCESSO em df_cabecalhos_global para criar mapeamento.")
            return map_id_info
        
        try:
            df_temp = df_cabecalhos_global.copy() 
            df_temp[col_id] = pd.to_numeric(df_temp[col_id], errors='coerce')
            df_temp.dropna(subset=[col_id], inplace=True) 
            
            for _, row in df_temp.iterrows():
                proc_id = int(row[col_id]) 
                nr_proc = str(row[col_nr]) if pd.notna(row[col_nr]) else 'N/A'
                juiz = str(row[col_juiz]) if pd.notna(row[col_juiz]) else self.config.DEFAULT_JUIZ_NAO_DESIGNADO
                map_id_info[proc_id] = {'nr_processo': nr_proc, 'juiz_auxiliar': juiz}
            return map_id_info
        except Exception as e:
            self.logger.error(f"Erro ao criar map ID -> NrProcesso e Juiz: {e}"); return {}

    def _criar_df_lista_processos(self, lista_ids, map_id_info_completo):
        col_id_nome = 'ID Processo'; col_nr_proc_nome = 'Nr Processo'; col_juiz_nome = self.config.COLUNA_JUIZ_AUXILIAR
        
        if not isinstance(lista_ids, list): lista_ids = []
        if not lista_ids: return pd.DataFrame(columns=[col_nr_proc_nome, col_id_nome, col_juiz_nome])
        
        lista_ids_int = []
        for id_proc in lista_ids:
            try:
                if pd.notna(id_proc): lista_ids_int.append(int(id_proc))
            except (ValueError, TypeError):
                self.logger.warning(f"Não foi possível converter o ID '{id_proc}' para int. Será ignorado.")
                continue


        data_for_df = []
        for proc_id_int in lista_ids_int:
            info = map_id_info_completo.get(proc_id_int, {})
            data_for_df.append({
                col_id_nome: proc_id_int,
                col_nr_proc_nome: info.get('nr_processo', 'N/A'),
                col_juiz_nome: info.get('juiz_auxiliar', self.config.DEFAULT_JUIZ_NAO_DESIGNADO)
            })
        
        df = pd.DataFrame(data_for_df)
        if df.empty: 
             return pd.DataFrame(columns=[col_nr_proc_nome, col_id_nome, col_juiz_nome])
        return df[[col_nr_proc_nome, col_id_nome, col_juiz_nome]]


    def _criar_df_sumario(self, resultados_por_escopo, escopos_calculados):
        sumario_data_list = []
        ordem_escopos = ['Global'] + sorted([e for e in escopos_calculados if e not in ['Global', self.config.DEFAULT_JUIZ_NAO_DESIGNADO]])
        if self.config.DEFAULT_JUIZ_NAO_DESIGNADO in escopos_calculados and self.config.DEFAULT_JUIZ_NAO_DESIGNADO not in ordem_escopos: 
            ordem_escopos.append(self.config.DEFAULT_JUIZ_NAO_DESIGNADO)
        
        for escopo in ordem_escopos:
            if escopo not in resultados_por_escopo: continue 
            
            resultados_escopo = resultados_por_escopo[escopo]
            res_m1 = resultados_escopo.get('meta1', {})
            res_m2 = resultados_escopo.get('meta2', {})
            res_m3 = resultados_escopo.get('meta3', {})

            sumario_data_list.append({
                'Escopo': escopo, 'Meta': 'Meta 1',
                'Indicador Base': 'P1.1 (Distr. Ano Meta)', 'Valor Base': res_m1.get('P1.1', 0), 
                'Indicador Meta': 'P1.2 (Todos Baixados Ano Meta)', 'Valor Meta': res_m1.get('P1.2', 0),
                'Percentual (%)': f"{res_m1.get('percentual', 0.0):.2f}",
                'Info Extra': f"P1.3 (Saldo Ano Anterior) = {res_m1.get('P1.3', 0)}"
            })
            sumario_data_list.append({
                'Escopo': escopo, 'Meta': 'Meta 2',
                'Indicador Base': f"P2.1 (Pend. Fim {self.config.ANO_META-1} de Autuados até {self.config.DATA_CORTE_META2.strftime('%d/%m/%Y')})", 'Valor Base': res_m2.get('P2.1', 0),
                'Indicador Meta': 'P2.2 (Resolvidos P2.1)', 'Valor Meta': res_m2.get('P2.2', 0),
                'Percentual (%)': f"{res_m2.get('percentual', 0.0):.2f}",
                'Info Extra': ""
            })
            sumario_data_list.append({
                'Escopo': escopo, 'Meta': 'Meta 3',
                'Indicador Base': f"P3.1 (Decid. {self.config.ANO_META})", 'Valor Base': res_m3.get('P3.1', 0),
                'Indicador Meta': 'P3.2 (Decid. Prazo)', 'Valor Meta': res_m3.get('P3.2', 0),
                'Percentual (%)': f"{res_m3.get('percentual', 0.0):.2f}",
                'Info Extra': f"Prazo <= {self.config.PRAZO_DIAS_META3} dias"
            })
        if not sumario_data_list: 
             return pd.DataFrame([{'Escopo': 'Nenhum dado', 'Meta': '', 'Indicador Base': '', 'Valor Base': '', 'Indicador Meta': '', 'Valor Meta': '', 'Percentual (%)': '', 'Info Extra': ''}])
        return pd.DataFrame(sumario_data_list)

    def _criar_map_tarefa_atual(self, df_tarefas_global): 
        self.logger.info("Pré-processando tarefas...")
        if df_tarefas_global is None or df_tarefas_global.empty: 
            self.logger.warning("DataFrame de Tarefas vazio ou não fornecido. Mapa de tarefas estará vazio.")
            return {}
        
        col_id=self.config.COLUNA_TAREFA_ID_PROCESSO; col_inicio=self.config.COLUNA_TAREFA_INICIO; col_fim=self.config.COLUNA_TAREFA_FIM
        col_fluxo=self.config.COLUNA_TAREFA_FLUXO; col_tarefa=self.config.COLUNA_TAREFA_NOME
        req_cols = [col_id, col_inicio, col_fim, col_fluxo, col_tarefa]
        if not all(c in df_tarefas_global.columns for c in req_cols): self.logger.error(f"ERRO: Faltam colunas em Tarefas: {req_cols}. Mapa de tarefas estará vazio."); return {}

        df = df_tarefas_global.copy()
        df[col_inicio] = pd.to_datetime(df[col_inicio], errors='coerce', dayfirst=False) 
        df[col_fim] = pd.to_datetime(df[col_fim], errors='coerce', dayfirst=False)
        df[col_id] = pd.to_numeric(df[col_id], errors='coerce').astype('Int64')
        df.dropna(subset=[col_id, col_inicio], inplace=True)

        abertas = df[df[col_fim].isnull()].copy()
        if abertas.empty: self.logger.info("Nenhuma tarefa aberta encontrada."); return {}

        abertas.sort_values(by=[col_id, col_inicio], ascending=[True, False], inplace=True)
        ultimas_abertas = abertas.drop_duplicates(subset=[col_id], keep='first')

        map_tarefas = {}
        for _, row in ultimas_abertas.iterrows():
            proc_id = row[col_id] 
            if pd.isna(proc_id): continue 
            fluxo_str = str(row[col_fluxo]) if pd.notna(row[col_fluxo]) else ""
            tarefa_str = str(row[col_tarefa]) if pd.notna(row[col_tarefa]) else ""
            map_tarefas[int(proc_id)] = f"{fluxo_str} > {tarefa_str}" 

        self.logger.info(f"Map ID->Tarefa Atual criado para {len(map_tarefas)} processos.")
        return map_tarefas

    def _criar_df_pendentes_com_tarefa(self, lista_ids, map_id_info_completo, map_id_tarefa):
        col_id_nome = 'ID Processo'; col_nr_proc_nome = 'Nr Processo'; col_juiz_nome = self.config.COLUNA_JUIZ_AUXILIAR
        col_tarefa_nome = 'Tarefa Atual'

        if not isinstance(lista_ids, list): lista_ids = []
        if not lista_ids: return pd.DataFrame(columns=[col_nr_proc_nome, col_id_nome, col_juiz_nome, col_tarefa_nome])

        lista_ids_int = []
        for id_proc in lista_ids:
            try:
                if pd.notna(id_proc): lista_ids_int.append(int(id_proc))
            except (ValueError, TypeError):
                self.logger.warning(f"Não foi possível converter o ID '{id_proc}' para int em pendentes com tarefa. Será ignorado.")
                continue

        data_for_df = []
        for proc_id_int in lista_ids_int:
            info_proc = map_id_info_completo.get(proc_id_int, {})
            data_for_df.append({
                col_id_nome: proc_id_int,
                col_nr_proc_nome: info_proc.get('nr_processo', 'N/A'),
                col_juiz_nome: info_proc.get('juiz_auxiliar', self.config.DEFAULT_JUIZ_NAO_DESIGNADO),
                col_tarefa_nome: map_id_tarefa.get(proc_id_int, 'Status Desconhecido')
            })
        
        df = pd.DataFrame(data_for_df)
        if df.empty:
            return pd.DataFrame(columns=[col_nr_proc_nome, col_id_nome, col_juiz_nome, col_tarefa_nome])
        return df[[col_nr_proc_nome, col_id_nome, col_juiz_nome, col_tarefa_nome]]

    def _criar_df_pendentes_prazo_meta3_com_tarefa(self, ids_relevantes_global, df_cabecalhos_global, df_movimentos_global, map_id_info_completo, map_id_tarefa):
        if ids_relevantes_global is None or ids_relevantes_global.empty or df_cabecalhos_global is None or df_movimentos_global is None: 
            return pd.DataFrame()
        self.logger.info("Identificando processos abertos e calculando prazo/tarefa Meta 3 (Global)...")
        col_id=self.config.COLUNA_ID_PROCESSO; col_aut=self.config.COLUNA_DATA_AUTUACAO
        
        df_cabecalhos_temp = df_cabecalhos_global.copy()
        if col_id not in df_cabecalhos_temp.columns or col_aut not in df_cabecalhos_temp.columns:
            self.logger.error(f"Coluna {col_id} ou {col_aut} não encontrada em df_cabecalhos_global para prazos M3.")
            return pd.DataFrame()

        df_cabecalhos_temp[col_id] = pd.to_numeric(df_cabecalhos_temp[col_id], errors='coerce')
        df_cabecalhos_temp.dropna(subset=[col_id], inplace=True)
        df_cabecalhos_temp[col_id] = df_cabecalhos_temp[col_id].astype(int)


        map_id_dt_aut = df_cabecalhos_temp.set_index(col_id)[col_aut].dropna().to_dict()
        
        processos_abertos_data = []; hoje = datetime.now(); prazo_m3 = self.config.PRAZO_DIAS_META3
        
        ids_relevantes_int_list = []
        if hasattr(ids_relevantes_global, 'tolist'):
            ids_relevantes_int_list = [int(i) for i in ids_relevantes_global.tolist() if pd.notna(i)]
        elif isinstance(ids_relevantes_global, list):
             ids_relevantes_int_list = [int(i) for i in ids_relevantes_global if pd.notna(i)]


        df_mov_rel_glob = df_movimentos_global[df_movimentos_global[col_id].isin(ids_relevantes_int_list)].copy()
        mov_rel_grouped = df_mov_rel_glob.groupby(col_id)

        for proc_id_int in ids_relevantes_int_list: 
            df_proc_mov = mov_rel_grouped.get_group(proc_id_int) if proc_id_int in mov_rel_grouped.groups else pd.DataFrame()
            status_atual = self._calculadora_m2_helper._get_terminal_status(df_proc_mov.copy(), hoje) 
            
            if status_atual == 'PENDING':
                dt_autuacao = map_id_dt_aut.get(proc_id_int)
                if not pd.isna(dt_autuacao):
                    info_proc = map_id_info_completo.get(proc_id_int, {})
                    nr_proc = info_proc.get('nr_processo', 'N/A')
                    juiz_aux = info_proc.get('juiz_auxiliar', self.config.DEFAULT_JUIZ_NAO_DESIGNADO)
                    tarefa_atual = map_id_tarefa.get(proc_id_int, 'Status Desconhecido')
                    dt_limite = dt_autuacao + timedelta(days=prazo_m3)
                    dias_rest = (dt_limite - hoje).days
                    processos_abertos_data.append({
                        'Nr Processo': nr_proc, 'ID Processo': proc_id_int, 
                        self.config.COLUNA_JUIZ_AUXILIAR: juiz_aux,
                        'Tarefa Atual': tarefa_atual, 'Data Autuação': dt_autuacao, 
                        f'Data Limite ({prazo_m3}d)': dt_limite, 'Dias Restantes': dias_rest
                    })
        if not processos_abertos_data: return pd.DataFrame()
        df_prazos = pd.DataFrame(processos_abertos_data)
        self.logger.info(f"Encontrados {len(df_prazos)} processos abertos (globais) para prazo Meta 3.")
        cols_ordem = ['Nr Processo', 'ID Processo', self.config.COLUNA_JUIZ_AUXILIAR, 'Tarefa Atual', 'Data Autuação', f'Data Limite ({prazo_m3}d)', 'Dias Restantes']
        if not all(c in df_prazos.columns for c in cols_ordem if c in df_prazos.columns): 
            self.logger.warning("Algumas colunas esperadas não foram geradas para df_prazos. A ordem pode estar incorreta.")
        else:
            df_prazos = df_prazos[cols_ordem] 
        
        if 'Data Autuação' in df_prazos.columns: 
            df_prazos['Data Autuação'] = pd.to_datetime(df_prazos['Data Autuação']).dt.strftime('%d/%m/%Y')
        if f'Data Limite ({prazo_m3}d)' in df_prazos.columns:
            df_prazos[f'Data Limite ({prazo_m3}d)'] = pd.to_datetime(df_prazos[f'Data Limite ({prazo_m3}d)']).dt.strftime('%d/%m/%Y')
        return df_prazos.sort_values(by=['Dias Restantes', self.config.COLUNA_JUIZ_AUXILIAR], ascending=[True, True])


    def salvar_relatorio(self, dados_analise, caminho_saida):
        try: import xlsxwriter
        except ImportError: self.logger.error("ERRO: 'xlsxwriter' não instalado."); return False
        if not dados_analise or 'resultados_por_escopo' not in dados_analise: self.logger.error("Dados inválidos p/ relatório."); return False
        
        resultados_por_escopo = dados_analise['resultados_por_escopo']
        df_cabecalhos_global = dados_analise.get('df_cabecalhos_global') 
        df_movimentos_global = dados_analise.get('df_movimentos_global')
        df_tarefas_global = dados_analise.get('df_tarefas_global')
        ids_relevantes_global = dados_analise.get('ids_relevantes_global')
        escopos_calculados = dados_analise.get('escopos_calculados', ['Global'])

        if df_cabecalhos_global is None or ids_relevantes_global is None or df_movimentos_global is None:
            self.logger.error("Dados globais insuficientes p/ detalhes (faltam DFs ou IDs globais)."); return False

        map_id_info_completo = self._criar_map_id_nrprocesso_e_juiz(df_cabecalhos_global)
        map_id_tarefa_atual = self._criar_map_tarefa_atual(df_tarefas_global) 

        if not map_id_info_completo: self.logger.warning("Mapeamento ID->Info Processo (Nr, Juiz) vazio.")

        self.logger.info(f"Gerando relatório final: {caminho_saida}")
        try:
            with pd.ExcelWriter(caminho_saida, engine='xlsxwriter', datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy') as writer:
                df_sumario = self._criar_df_sumario(resultados_por_escopo, escopos_calculados)
                df_sumario.to_excel(writer, sheet_name='Sumário_Metas', index=False)
                self.logger.info("Aba Sumário (por Escopo) criada.")

                res_m1_global = resultados_por_escopo.get('Global', {}).get('meta1', {})
                ids_pend_m1_global = list(set(res_m1_global.get('ids_P1.3', [])) - set(res_m1_global.get('ids_P1.2', [])))
                df_pend_m1_global = self._criar_df_pendentes_com_tarefa(ids_pend_m1_global, map_id_info_completo, map_id_tarefa_atual)
                if not df_pend_m1_global.empty: df_pend_m1_global.to_excel(writer, sheet_name='Acao_M1_Pend_Global', index=False)

                res_m2_global = resultados_por_escopo.get('Global', {}).get('meta2', {})
                ids_pend_m2_global = list(set(res_m2_global.get('ids_P2.1', [])) - set(res_m2_global.get('ids_P2.2', [])))
                df_pend_m2_global = self._criar_df_pendentes_com_tarefa(ids_pend_m2_global, map_id_info_completo, map_id_tarefa_atual)
                if not df_pend_m2_global.empty: df_pend_m2_global.to_excel(writer, sheet_name='Acao_M2_Pend_Global', index=False)

                df_prazos_m3_global = self._criar_df_pendentes_prazo_meta3_com_tarefa(ids_relevantes_global, df_cabecalhos_global, df_movimentos_global, map_id_info_completo, map_id_tarefa_atual)
                if not df_prazos_m3_global.empty: df_prazos_m3_global.to_excel(writer, sheet_name='Acao_M3_Pend_Prazo_Global', index=False)
                self.logger.info("Abas de Ação Globais criadas.")

                juizes_para_abas_acao = sorted([esc for esc in escopos_calculados if esc not in ['Global', self.config.DEFAULT_JUIZ_NAO_DESIGNADO]])
                if self.config.DEFAULT_JUIZ_NAO_DESIGNADO in escopos_calculados and self.config.DEFAULT_JUIZ_NAO_DESIGNADO not in juizes_para_abas_acao:
                    juizes_para_abas_acao.append(self.config.DEFAULT_JUIZ_NAO_DESIGNADO)

                for juiz_escopo in juizes_para_abas_acao:
                    nome_aba_juiz_raw = str(juiz_escopo).replace(" ", "_").replace("/", "-").replace("\\", "-")
                    # <<< CORREÇÃO Limite Nome Aba Excel >>>
                    nome_aba_juiz = "".join(c for c in nome_aba_juiz_raw if c.isalnum() or c in ['_','-'])[:18] 
                    
                    resultados_juiz = resultados_por_escopo.get(juiz_escopo, {})
                    
                    res_m1_juiz = resultados_juiz.get('meta1', {})
                    ids_pend_m1_juiz = list(set(res_m1_juiz.get('ids_P1.3', [])) - set(res_m1_juiz.get('ids_P1.2', [])))
                    df_pend_m1_juiz = self._criar_df_pendentes_com_tarefa(ids_pend_m1_juiz, map_id_info_completo, map_id_tarefa_atual)
                    if not df_pend_m1_juiz.empty: df_pend_m1_juiz.to_excel(writer, sheet_name=f"Acao_M1_Pend_{nome_aba_juiz}", index=False)

                    res_m2_juiz = resultados_juiz.get('meta2', {})
                    ids_pend_m2_juiz = list(set(res_m2_juiz.get('ids_P2.1', [])) - set(res_m2_juiz.get('ids_P2.2', [])))
                    df_pend_m2_juiz = self._criar_df_pendentes_com_tarefa(ids_pend_m2_juiz, map_id_info_completo, map_id_tarefa_atual)
                    if not df_pend_m2_juiz.empty: df_pend_m2_juiz.to_excel(writer, sheet_name=f"Acao_M2_Pend_{nome_aba_juiz}", index=False)
                
                self.logger.info("Abas de Ação por Juiz/Categoria criadas.")


                if ids_relevantes_global is not None and not ids_relevantes_global.empty:
                    df_ids_rel_global = self._criar_df_lista_processos(ids_relevantes_global.tolist(), map_id_info_completo)
                    if not df_ids_rel_global.empty: df_ids_rel_global.to_excel(writer, sheet_name='Detalhe_IDs_Relev_Global', index=False)

                mapa_abas_indicadores = {
                    'P1.1': ('meta1', 'ids_P1.1'), 'P1.2': ('meta1', 'ids_P1.2'), 'P1.3': ('meta1', 'ids_P1.3'),
                    'P2.1': ('meta2', 'ids_P2.1'), 'P2.2': ('meta2', 'ids_P2.2'),
                    'P3.1': ('meta3', 'details_P3.1'), 'P3.2': ('meta3', 'details_P3.2')
                }
                
                resultados_globais = resultados_por_escopo.get('Global', {})
                if resultados_globais: 
                    for nome_aba_indicador, (chave_meta, chave_dados) in mapa_abas_indicadores.items():
                        lista_ids_indicador = []; dados_meta_global = resultados_globais.get(chave_meta)
                        if dados_meta_global and chave_dados in dados_meta_global:
                            dados = dados_meta_global[chave_dados]
                            if isinstance(dados, list): lista_ids_indicador = dados
                            elif isinstance(dados, dict): lista_ids_indicador = list(dados.keys())
                        
                        df_indicador = self._criar_df_lista_processos(lista_ids_indicador, map_id_info_completo)
                        if not df_indicador.empty: 
                            df_indicador.to_excel(writer, sheet_name=f"Detalhe_{nome_aba_indicador}_Global", index=False)
                self.logger.info("Abas de Detalhe por Indicador (Global) criadas.")

            self.logger.info(f"Relatório salvo: {caminho_saida}"); return True
        except xlsxwriter.exceptions.InvalidWorksheetName as e_iwsn: 
            self.logger.error(f"ERRO DE NOME DE ABA INVÁLIDO ao gerar relatório '{caminho_saida}': {e_iwsn}")
            self.logger.error(traceback.format_exc()); return False
        except Exception as e:
            self.logger.error(f"ERRO GERAL ao gerar/salvar relatório '{caminho_saida}': {e}")
            self.logger.error(traceback.format_exc()); return False

# ==========================
# Classe da Interface Gráfica (GUI)
# (Sem alterações nesta classe em relação à v2.8)
# ==========================
class MonitorMetasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Monitor de Metas Corregedorias {ConfiguracaoMetas.ANO_META} (v2.9 M1 Fix)") # Atualiza versão no título
        self.root.geometry("850x750") 
        self.caminho_cabecalhos = tk.StringVar()
        self.caminho_movimentos = tk.StringVar()
        self.caminho_tarefas = tk.StringVar()
        self.caminho_saida = tk.StringVar()
        self.analise_em_andamento = False
        style = ttk.Style(); 
        try:
            if os.name == 'nt': style.theme_use('vista')
            else: style.theme_use('clam') 
        except tk.TclError: 
            self.log("Tema 'vista' ou 'clam' não disponível, usando default.", "WARNING")
            pass 
        self.criar_interface()

    def log(self, mensagem, nivel="INFO"):
        try:
            timestamp = datetime.now().strftime('%H:%M:%S'); log_entry = f"[{timestamp}] [{nivel}] {mensagem}\n"
            def update_log_gui(): 
                if hasattr(self, 'log_area') and self.log_area.winfo_exists():
                    self.log_area.configure(state=tk.NORMAL); self.log_area.insert(tk.END, log_entry); self.log_area.see(tk.END); self.log_area.configure(state=tk.DISABLED)
            if self.root and hasattr(self.root, 'after'): 
                 self.root.after(0, update_log_gui)
            else: 
                print(log_entry.strip()) 
            
            if nivel == "INFO": logging.info(mensagem)
            elif nivel == "WARNING": logging.warning(mensagem)
            elif nivel == "ERROR": logging.error(mensagem)
            elif nivel == "DEBUG": logging.debug(mensagem)

        except Exception as e: 
            print(f"Erro no log da GUI: {e}") 

    def criar_interface(self):
        main_frame = ttk.Frame(self.root, padding="10"); main_frame.pack(fill=tk.BOTH, expand=True)
        title_frame = ttk.Frame(main_frame); title_frame.pack(fill=tk.X, pady=5)
        ttk.Label(title_frame, text=f"Monitor Metas Corregedorias {ConfiguracaoMetas.ANO_META} (Global e Juízes)", font=("Helvetica", 16, "bold")).pack()
        ttk.Label(title_frame, text="Autor: Cristhiano Leite dos Santos | Versão: 2.9", font=("Helvetica", 9)).pack() # Atualiza versão

        file_frame = ttk.LabelFrame(main_frame, text="Arquivos de Dados", padding="10"); file_frame.pack(fill=tk.X, pady=10, padx=5)
        ttk.Label(file_frame, text="Cabeçalhos:").grid(row=0, column=0, sticky=tk.W, pady=3, padx=5)
        ttk.Entry(file_frame, textvariable=self.caminho_cabecalhos, width=70).grid(row=0, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(file_frame, text="Procurar...", command=self.selecionar_cabecalhos).grid(row=0, column=2, padx=5, pady=3)
        ttk.Label(file_frame, text="Movimentos:").grid(row=1, column=0, sticky=tk.W, pady=3, padx=5)
        ttk.Entry(file_frame, textvariable=self.caminho_movimentos, width=70).grid(row=1, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(file_frame, text="Procurar...", command=self.selecionar_movimentos).grid(row=1, column=2, padx=5, pady=3)
        ttk.Label(file_frame, text="Tarefas:").grid(row=2, column=0, sticky=tk.W, pady=3, padx=5)
        ttk.Entry(file_frame, textvariable=self.caminho_tarefas, width=70).grid(row=2, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(file_frame, text="Procurar...", command=self.selecionar_tarefas).grid(row=2, column=2, padx=5, pady=3)
        ttk.Label(file_frame, text="Salvar Relatório:").grid(row=3, column=0, sticky=tk.W, pady=3, padx=5)
        ttk.Entry(file_frame, textvariable=self.caminho_saida, width=70).grid(row=3, column=1, padx=5, pady=3, sticky="ew")
        ttk.Button(file_frame, text="Salvar Como...", command=self.selecionar_saida).grid(row=3, column=2, padx=5, pady=3)
        file_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(main_frame); button_frame.pack(pady=10)
        self.btn_executar = ttk.Button(button_frame, text="Executar Análise Completa", command=self.executar_analise_gui, width=25); self.btn_executar.pack(side=tk.LEFT, padx=10)
        self.btn_limpar = ttk.Button(button_frame, text="Limpar Campos e Log", command=self.limpar_campos, width=20); self.btn_limpar.pack(side=tk.LEFT, padx=10)

        self.progress_bar = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='indeterminate'); self.progress_bar.pack(fill=tk.X, pady=5, padx=5)
        log_frame = ttk.LabelFrame(main_frame, text="Log de Execução", padding="5"); log_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=5)
        self.log_area = tk.Text(log_frame, wrap=tk.WORD, height=18, font=("Consolas", 9), state=tk.DISABLED, bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_area.yview); self.log_area.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y); self.log_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def selecionar_arquivo(self, title, filetypes, variable):
        filename = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if filename: variable.set(filename); self.log(f"Arquivo selecionado: {os.path.basename(filename)}")

    def selecionar_cabecalhos(self): self.selecionar_arquivo("Selecione Cabeçalhos",(("CSV", "*.csv"),("Excel", "*.xlsx *.xls"),("Todos", "*.*")), self.caminho_cabecalhos)
    def selecionar_movimentos(self): self.selecionar_arquivo("Selecione Movimentos",(("CSV", "*.csv"),("Excel", "*.xlsx *.xls"),("Todos", "*.*")), self.caminho_movimentos)
    def selecionar_tarefas(self): self.selecionar_arquivo("Selecione Tarefas (Opcional)", (("CSV", "*.csv"),("Excel", "*.xlsx *.xls"),("Todos", "*.*")), self.caminho_tarefas) 
    def selecionar_saida(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S"); dfn = f"Relatorio_Metas_Corregedorias_{ConfiguracaoMetas.ANO_META}_{ts}.xlsx" 
        filename = filedialog.asksaveasfilename(title="Salvar Relatório", defaultextension=".xlsx", initialfile=dfn, filetypes=(("Excel", "*.xlsx"),("Todos", "*.*")))
        if filename: self.caminho_saida.set(filename); self.log(f"Arquivo de saída: {os.path.basename(filename)}")

    def limpar_campos(self):
        self.caminho_cabecalhos.set(""); self.caminho_movimentos.set(""); self.caminho_tarefas.set(""); self.caminho_saida.set("")
        if hasattr(self, 'log_area') and self.log_area.winfo_exists(): self.log_area.configure(state=tk.NORMAL); self.log_area.delete('1.0', tk.END); self.log_area.configure(state=tk.DISABLED)
        self.log("Campos e log da GUI limpos.")

    def executar_analise_gui(self):
        if self.analise_em_andamento: messagebox.showwarning("Atenção", "Análise já em andamento."); return
        cabecalho=self.caminho_cabecalhos.get(); movimentos=self.caminho_movimentos.get(); tarefas=self.caminho_tarefas.get(); saida=self.caminho_saida.get()

        if not cabecalho or not os.path.exists(cabecalho): messagebox.showerror("Erro de Arquivo", "Arquivo de Cabeçalhos é obrigatório e não foi encontrado."); return
        if not movimentos or not os.path.exists(movimentos): messagebox.showerror("Erro de Arquivo", "Arquivo de Movimentos é obrigatório e não foi encontrado."); return
        if tarefas and not os.path.exists(tarefas): messagebox.showerror("Erro de Arquivo", "Arquivo de Tarefas fornecido não foi encontrado."); return 

        if not saida:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S"); dfn = f"Relatorio_Metas_Corregedorias_{ConfiguracaoMetas.ANO_META}_{ts}.xlsx"
            try: d_script = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
            except NameError: d_script = os.getcwd()
            saida = os.path.join(d_script, dfn); self.caminho_saida.set(saida); self.log(f"Arquivo de saída padrão definido: {os.path.basename(saida)}")
        elif not saida.lower().endswith(".xlsx"): messagebox.showerror("Erro de Arquivo", "O nome do arquivo de saída deve terminar com .xlsx"); return

        self.analise_em_andamento = True; self.btn_executar.config(state=tk.DISABLED); self.btn_limpar.config(state=tk.DISABLED)
        self.progress_bar.start(10); self.log("="*40); self.log("INICIANDO ANÁLISE COMPLETA..."); 
        self.log(f"Cabeçalhos: {os.path.basename(cabecalho)}"); self.log(f"Movimentos: {os.path.basename(movimentos)}")
        if tarefas: self.log(f"Tarefas: {os.path.basename(tarefas)}")
        else: self.log("Tarefas: Nenhum arquivo selecionado (opcional).")
        self.log(f"Saída: {os.path.basename(saida)}"); self.log("="*40)
        
        backend_logger = logging.getLogger()
        threading.Thread(target=self._processar_analise_thread, args=(cabecalho, movimentos, tarefas, saida, backend_logger), daemon=True).start()

    def _finalizar_analise_gui(self, sucesso=True, caminho_saida=None):
        self.progress_bar.stop(); self.btn_executar.config(state=tk.NORMAL); self.btn_limpar.config(state=tk.NORMAL); self.analise_em_andamento = False
        self.log("="*40); self.log("ANÁLISE FINALIZADA."); self.log("="*40)
        if sucesso and caminho_saida: 
            messagebox.showinfo("Concluído com Sucesso", f"Análise concluída!\nRelatório salvo em:\n{caminho_saida}")
        elif not sucesso and caminho_saida: 
             messagebox.showwarning("Concluído com Erros", f"Análise concluída com erros. Verifique o log.\nRelatório (pode estar incompleto) salvo em:\n{caminho_saida}")
        elif not sucesso and not caminho_saida: 
            messagebox.showerror("Erro na Análise", "A análise falhou antes de gerar o relatório. Verifique o log.")


    def _processar_analise_thread(self, caminho_cabecalho, caminho_movimentos, caminho_tarefas, caminho_saida, backend_logger):
        sucesso_geral = False
        caminho_saida_final = None 
        try:
            self.log("Backend: Iniciando AnalisadorMetas...") 
            analisador = AnalisadorMetas(config=ConfiguracaoMetas(), logger=backend_logger) 
            dados_completos_analise = analisador.executar_analise(caminho_cabecalho, caminho_movimentos, caminho_tarefas)

            if dados_completos_analise and dados_completos_analise.get('resultados_por_escopo'):
                falhas_reportadas = False
                for escopo, resultados_escopo in dados_completos_analise['resultados_por_escopo'].items():
                    for nome_meta, res_meta in resultados_escopo.items():
                        if not res_meta or not isinstance(res_meta, dict) or res_meta.get('percentual') is None: 
                            self.log(f"ATENÇÃO: Falha ou dados incompletos para {nome_meta.upper()} no escopo '{escopo}'.", "WARNING")
                            falhas_reportadas = True
                
                self.log("Backend: Gerando Relatório Excel...", "INFO")
                gerador = GeradorRelatorio(config=ConfiguracaoMetas(), logger=backend_logger)
                sucesso_salvar = gerador.salvar_relatorio(dados_completos_analise, caminho_saida)
                caminho_saida_final = caminho_saida 
                
                if sucesso_salvar: 
                    self.log("Backend: Relatório gerado com sucesso.", "INFO"); sucesso_geral = True
                else: 
                    msg_erro = "Backend: Falha crítica ao gerar ou salvar o relatório Excel."; self.log(msg_erro, "ERROR")
                    if hasattr(self, 'root') and self.root: self.root.after(0, lambda: messagebox.showerror("Erro Crítico ao Salvar", msg_erro))
            else: 
                msg_erro = "Backend: Análise não retornou dados ou resultados_por_escopo está vazio."; self.log(msg_erro, "ERROR")
                if hasattr(self, 'root') and self.root: self.root.after(0, lambda: messagebox.showerror("Erro na Análise", msg_erro))
        
        except Exception as e:
            error_msg = f"ERRO GERAL NO PROCESSAMENTO DA ANÁLISE: {e}"; 
            detail_msg = traceback.format_exc()
            self.log(error_msg, "ERROR"); 
            backend_logger.error(error_msg + "\n" + detail_msg) 
            if hasattr(self, 'root') and self.root: 
                self.root.after(0, lambda: messagebox.showerror("Erro Inesperado no Processamento", error_msg))
        finally:
            if hasattr(self, 'root') and self.root: 
                self.root.after(0, lambda s=sucesso_geral, cs=caminho_saida_final: self._finalizar_analise_gui(sucesso=s, caminho_saida=cs))

# ==========================
# Bloco Principal (Execução da GUI)
# ==========================
if __name__ == "__main__":
    log_stream = io.StringIO()
    logging.basicConfig(level=logging.INFO, 
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        handlers=[logging.StreamHandler(log_stream), logging.StreamHandler()]) 
    logger_main = logging.getLogger(__name__)

    try: import xlsxwriter; logger_main.info("'xlsxwriter' encontrado.")
    except ImportError: 
        logger_main.warning("Biblioteca 'xlsxwriter' não encontrada. Exportação Excel falhará. Instale: pip install xlsxwriter")
    try: import requests; logger_main.info("'requests' encontrado.")
    except ImportError:
        logger_main.warning("Biblioteca 'requests' não encontrada. Busca de dados da API falhará. Instale: pip install requests")

    root = tk.Tk()
    app = MonitorMetasGUI(root)
    root.mainloop()