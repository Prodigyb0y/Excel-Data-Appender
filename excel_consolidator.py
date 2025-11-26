import os
import logging
import xlwings as xw
from pathlib import Path

# Configuração de Logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("ExcelBot")

def verificar_arquivos(origem: str, destino: str) -> bool:
    """Verifica se os arquivos de entrada existem."""
    if not os.path.exists(origem):
        logger.error(f"Arquivo de origem não encontrado: {origem}")
        return False
    if not os.path.exists(destino):
        logger.error(f"Arquivo de destino não encontrado: {destino}")
        return False
    return True

def consolidar_dados_excel(
    path_origem: str, 
    sheet_origem: str, 
    path_destino: str, 
    sheet_destino: str
):
    """
    Copia dados de uma planilha de origem e anexa ao final de uma planilha de destino.
    Utiliza xlwings para automação via COM interface.
    """
    
    # Validação
    if not verificar_arquivos(path_origem, path_destino):
        return

    app = None
    try:
        # Inicia uma instância do Excel (invisível para ser mais rápido/seguro)
        # Nota: Com xlwings, às vezes é melhor usar a instância ativa ou app.visible=False
        app = xw.App(visible=False)
        
        logger.info(f"Abrindo origem: {Path(path_origem).name}")
        wb_origem = app.books.open(path_origem)
        ws_origem = wb_origem.sheets[sheet_origem]

        logger.info(f"Abrindo destino: {Path(path_destino).name}")
        wb_destino = app.books.open(path_destino)
        ws_destino = wb_destino.sheets[sheet_destino]

        # 1. Mapear última linha da origem (Baseado na Coluna B)
        # Equivalente ao Ctrl+Setinha pra cima
        last_row_origem = ws_origem.range(f"B{ws_origem.cells.last_cell.row}").end('up').row
        
        if last_row_origem < 3:
            logger.warning("Não há dados suficientes para copiar na origem (começa na linha 3).")
            return

        # 2. Mapear onde colar no destino (Baseado na Coluna J - índice 10)
        # Procura a última célula preenchida na coluna J e soma 1
        last_row_destino = ws_destino.range(ws_destino.cells.last_cell.row, 10).end('up').row
        target_row = last_row_destino + 1

        # 3. Operação de Cópia
        intervalo_origem = ws_origem.range(f"B3:H{last_row_origem}")
        celula_destino = ws_destino.range((target_row, 10)) # Coluna J é a 10ª

        logger.info(f"Copiando {last_row_origem - 2} linhas para o destino a partir da linha {target_row}...")
        
        intervalo_origem.copy()
        celula_destino.paste() # Paste simples (pode usar paste='values' se quiser só valores)

        # Salvar e Fechar
        wb_destino.save()
        logger.info("✅ Dados consolidados com sucesso.")

    except Exception as e:
        logger.exception(f"❌ Erro crítico durante a execução: {e}")
    
    finally:
        # Garante que o Excel feche e libere a memória
        if app:
            try:
                # Fecha os livros sem salvar o de origem (para não alterar nada lá acidentalmente)
                # O destino já foi salvo acima se tudo deu certo
                app.quit()
            except:
                pass
        logger.info("Processo finalizado.")

if __name__ == "__main__":
    # --- ÁREA DE CONFIGURAÇÃO ---
    # Dica: Use raw strings (r"") para caminhos Windows
    ARQUIVO_ORIGEM = r"C:\Caminho\Para\Sua\Pasta\Ordens_Encerradas.xlsx"
    ABA_ORIGEM = "Ente_Seg"

    ARQUIVO_DESTINO = r"C:\Caminho\Para\Sua\Pasta\Analise_Geral.xlsm"
    ABA_DESTINO = "Geral"

    consolidar_dados_excel(ARQUIVO_ORIGEM, ABA_ORIGEM, ARQUIVO_DESTINO, ABA_DESTINO)
