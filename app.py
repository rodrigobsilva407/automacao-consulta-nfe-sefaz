from __future__ import annotations

import os
import sys
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QObject, Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QTextCursor
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

STYLESHEET = """
QMainWindow, QWidget#central {
    background-color: #0d1117;
}
QWidget {
    font-family: 'Segoe UI', Arial, sans-serif;
    color: #e6edf3;
}
QFrame#card {
    background-color: #161b22;
    border: 1px solid #30363d;
    border-radius: 8px;
}
QLabel#title {
    color: #58a6ff;
    font-size: 26px;
    font-weight: bold;
}
QLabel#subtitle {
    color: #8b949e;
    font-size: 13px;
}
QLabel#section {
    color: #8b949e;
    font-size: 11px;
    font-weight: bold;
}
QLabel#stat-value {
    color: #58a6ff;
    font-size: 22px;
    font-weight: bold;
}
QLabel#stat-label {
    color: #8b949e;
    font-size: 11px;
}
QLineEdit {
    background-color: #0d1117;
    border: 1px solid #30363d;
    border-radius: 6px;
    padding: 10px 14px;
    color: #e6edf3;
    font-size: 14px;
}
QLineEdit:focus {
    border-color: #58a6ff;
}
QSpinBox {
    background-color: #0d1117;
    border: 1px solid #30363d;
    border-radius: 6px;
    padding: 6px 10px;
    color: #e6edf3;
    font-size: 13px;
}
QSpinBox:focus {
    border-color: #58a6ff;
}
QSpinBox::up-button, QSpinBox::down-button {
    background-color: #21262d;
    border: none;
    width: 18px;
}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background-color: #30363d;
}
QPushButton#browse {
    background-color: #21262d;
    border: 1px solid #30363d;
    border-radius: 6px;
    padding: 10px 18px;
    color: #e6edf3;
    font-size: 14px;
}
QPushButton#browse:hover {
    background-color: #30363d;
    border-color: #58a6ff;
}
QPushButton#run {
    background-color: #1F4E78;
    border: none;
    border-radius: 8px;
    padding: 14px 36px;
    color: white;
    font-size: 15px;
    font-weight: bold;
    min-height: 50px;
}
QPushButton#run:hover {
    background-color: #2563a8;
}
QPushButton#run:disabled {
    background-color: #21262d;
    color: #8b949e;
}
QPushButton#cancel {
    background-color: #6e2020;
    border: none;
    border-radius: 8px;
    padding: 14px 36px;
    color: white;
    font-size: 15px;
    font-weight: bold;
    min-height: 50px;
}
QPushButton#cancel:hover {
    background-color: #a83232;
}
QPushButton#cancel:disabled {
    background-color: #21262d;
    color: #8b949e;
}
QPushButton#open {
    background-color: #238636;
    border: none;
    border-radius: 8px;
    padding: 14px 36px;
    color: white;
    font-size: 15px;
    font-weight: bold;
    min-height: 50px;
}
QPushButton#open:hover {
    background-color: #2ea043;
}
QTextEdit#log {
    background-color: #0d1117;
    border: none;
    padding: 10px;
    color: #7ee787;
    font-family: 'Cascadia Code', 'Consolas', monospace;
    font-size: 13px;
}
QProgressBar {
    background-color: #21262d;
    border: none;
    border-radius: 4px;
    height: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #1F4E78;
    border-radius: 4px;
}
QTabWidget::pane {
    background-color: #161b22;
    border: 1px solid #30363d;
    border-radius: 0 8px 8px 8px;
}
QTabBar::tab {
    background-color: #0d1117;
    color: #8b949e;
    border: 1px solid #30363d;
    border-bottom: none;
    border-radius: 6px 6px 0 0;
    padding: 6px 18px;
    margin-right: 2px;
    font-size: 12px;
}
QTabBar::tab:selected {
    background-color: #161b22;
    color: #e6edf3;
    border-bottom-color: #161b22;
}
QTabBar::tab:hover:!selected {
    background-color: #21262d;
    color: #e6edf3;
}
QTableWidget {
    background-color: #0d1117;
    border: none;
    gridline-color: #30363d;
    color: #e6edf3;
    font-size: 13px;
}
QTableWidget::item {
    padding: 6px 10px;
    border-bottom: 1px solid #21262d;
}
QTableWidget::item:selected {
    background-color: #1F4E78;
}
QHeaderView::section {
    background-color: #161b22;
    color: #8b949e;
    border: none;
    border-bottom: 1px solid #30363d;
    padding: 6px 10px;
    font-size: 11px;
    font-weight: bold;
}
QStatusBar {
    background-color: #161b22;
    color: #8b949e;
    font-size: 11px;
    border-top: 1px solid #30363d;
}
QScrollBar:vertical {
    background: #161b22;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #30363d;
    border-radius: 4px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background: #58a6ff;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""


class LogStream(QObject):
    message = pyqtSignal(str)

    def write(self, text: str) -> None:
        text = text.strip()
        if text:
            self.message.emit(text)

    def flush(self) -> None:
        pass


class Worker(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal(bool, str)
    summary_signal = pyqtSignal(dict)

    def __init__(self, input_file: str, output_file: str, stop_event: threading.Event) -> None:
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self._stop_event = stop_event

    def run(self) -> None:
        stream = LogStream()
        stream.message.connect(self.log_signal)
        original_stdout = sys.stdout
        sys.stdout = stream  # type: ignore[assignment]

        try:
            os.chdir(Path(self.input_file).parent)

            import main as m

            chaves = m.carregar_chaves()
            self.log_signal.emit(f"Consultando {len(chaves)} chave(s)...")

            def on_progress(done: int, total: int) -> None:
                self.progress_signal.emit(done, total)

            notas, lancamentos, itens = m.consultar_notas_em_paralelo(
                chaves, on_progress, self._stop_event
            )

            if self._stop_event.is_set():
                self.finished_signal.emit(False, "Cancelado pelo usuário.")
                return

            m.gerar_excel_profissional(notas, lancamentos, itens, self.output_file)

            total_sucesso = sum(1 for n in notas if n.get("status_consulta") == "SUCESSO")
            total_erros = len(notas) - total_sucesso
            valor_total = sum(float(n.get("valor_nota") or 0) for n in notas)
            self.summary_signal.emit({
                "total_chaves": len(chaves),
                "total_notas": len(notas),
                "sucesso": total_sucesso,
                "erros": total_erros,
                "lancamentos": len(lancamentos),
                "itens": len(itens),
                "valor_total": valor_total,
            })

            self.finished_signal.emit(True, self.output_file)

        except Exception as e:
            self.finished_signal.emit(False, str(e))
        finally:
            sys.stdout = original_stdout


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.worker: Worker | None = None
        self._output_path: str = ""
        self._stop_event = threading.Event()
        self._start_time: datetime | None = None
        self._elapsed_timer = QTimer(self)
        self._elapsed_timer.setInterval(1000)
        self._elapsed_timer.timeout.connect(self._tick_elapsed)
        self._setup_ui()
        sb = self.statusBar()
        self._status_bar: QStatusBar = sb if sb is not None else QStatusBar(self)
        self._status_bar.showMessage("Pronto")

    # ── UI Setup ─────────────────────────────────────────────────────

    def _setup_ui(self) -> None:
        self.setWindowTitle("SITRAM — Relatório de NF-e")
        self.setMinimumSize(900, 700)
        self.resize(1100, 820)

        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("QScrollArea { background-color: #0d1117; border: none; }")
        outer.addWidget(scroll)

        scroll_content = QWidget()
        scroll_content.setObjectName("central")
        scroll.setWidget(scroll_content)
        root = QVBoxLayout(scroll_content)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(12)

        # ── Header ──────────────────────────────────────────────────
        header = QFrame()
        header.setObjectName("card")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 12, 20, 12)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title = QLabel("SITRAM")
        title.setObjectName("title")
        subtitle = QLabel("Portal SEFAZ-CE  ·  Relatório de NF-e")
        subtitle.setObjectName("subtitle")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        header_layout.addLayout(title_col)
        header_layout.addStretch()

        badge = QLabel("Beach Park")
        badge.setStyleSheet(
            "background-color: #1F4E78; color: white; border-radius: 12px;"
            "padding: 4px 14px; font-size: 11px; font-weight: bold;"
        )
        header_layout.addWidget(badge)
        root.addWidget(header)

        # ── Arquivo de Entrada ───────────────────────────────────────
        input_card = QFrame()
        input_card.setObjectName("card")
        input_layout = QVBoxLayout(input_card)
        input_layout.setContentsMargins(20, 12, 20, 12)
        input_layout.setSpacing(6)
        _lbl_entrada = QLabel("ARQUIVO DE ENTRADA")
        _lbl_entrada.setObjectName("section")
        input_layout.addWidget(_lbl_entrada)

        input_row = QHBoxLayout()
        input_row.setSpacing(8)
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Selecione o arquivo chaves.xlsx...")
        browse_input = QPushButton("Selecionar")
        browse_input.setObjectName("browse")
        browse_input.setFixedWidth(100)
        browse_input.clicked.connect(self._browse_input)
        input_row.addWidget(self.input_edit)
        input_row.addWidget(browse_input)
        input_layout.addLayout(input_row)
        root.addWidget(input_card)

        # ── Arquivo de Saída ─────────────────────────────────────────
        output_card = QFrame()
        output_card.setObjectName("card")
        output_layout = QVBoxLayout(output_card)
        output_layout.setContentsMargins(20, 12, 20, 12)
        output_layout.setSpacing(6)
        _lbl_saida = QLabel("ARQUIVO DE SAÍDA")
        _lbl_saida.setObjectName("section")
        output_layout.addWidget(_lbl_saida)

        output_row = QHBoxLayout()
        output_row.setSpacing(8)
        self.output_edit = QLineEdit("Relatorio_SITRAM.xlsx")
        self.output_edit.setPlaceholderText("Nome do relatório de saída...")
        browse_output = QPushButton("Selecionar")
        browse_output.setObjectName("browse")
        browse_output.setFixedWidth(100)
        browse_output.clicked.connect(self._browse_output)
        output_row.addWidget(self.output_edit)
        output_row.addWidget(browse_output)
        output_layout.addLayout(output_row)
        root.addWidget(output_card)

        # ── Botões ───────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        self.run_btn = QPushButton("▶   Gerar Relatório")
        self.run_btn.setObjectName("run")
        self.run_btn.clicked.connect(self._start)
        self.cancel_btn = QPushButton("✕   Cancelar")
        self.cancel_btn.setObjectName("cancel")
        self.cancel_btn.clicked.connect(self._cancel)
        self.cancel_btn.setEnabled(False)
        self.open_btn = QPushButton("📂   Abrir Relatório")
        self.open_btn.setObjectName("open")
        self.open_btn.clicked.connect(self._open_output)
        self.open_btn.setVisible(False)
        btn_row.addWidget(self.run_btn)
        btn_row.addWidget(self.cancel_btn)
        btn_row.addWidget(self.open_btn)
        root.addLayout(btn_row)

        # ── Progresso ────────────────────────────────────────────────
        progress_card = QFrame()
        progress_card.setObjectName("card")
        progress_layout = QVBoxLayout(progress_card)
        progress_layout.setContentsMargins(20, 12, 20, 12)
        progress_layout.setSpacing(6)

        progress_top = QHBoxLayout()
        _lbl_prog = QLabel("PROGRESSO")
        _lbl_prog.setObjectName("section")
        progress_top.addWidget(_lbl_prog)
        self.elapsed_label = QLabel("")
        self.elapsed_label.setStyleSheet("color: #58a6ff; font-size: 12px;")
        progress_top.addWidget(self.elapsed_label)
        progress_top.addStretch()
        self.progress_label = QLabel("Aguardando...")
        self.progress_label.setStyleSheet("color: #8b949e; font-size: 12px;")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        progress_top.addWidget(self.progress_label)
        progress_layout.addLayout(progress_top)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        progress_layout.addWidget(self.progress_bar)
        root.addWidget(progress_card)

        # ── Abas: Log / Resumo / Configurações ───────────────────────
        self.tabs = QTabWidget()
        self.tabs.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # Tab Log
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(12, 10, 12, 10)
        log_layout.setSpacing(6)

        log_top = QHBoxLayout()
        log_top.addStretch()
        clear_btn = QPushButton("Limpar")
        clear_btn.setFixedWidth(60)
        clear_btn.setStyleSheet(
            "background: transparent; color: #8b949e; border: none;"
            "font-size: 11px; padding: 0;"
        )
        clear_btn.clicked.connect(lambda: self.log_area.clear())
        log_top.addWidget(clear_btn)
        log_layout.addLayout(log_top)

        self.log_area = QTextEdit()
        self.log_area.setObjectName("log")
        self.log_area.setReadOnly(True)
        log_layout.addWidget(self.log_area)
        self.tabs.addTab(log_widget, "📋  Log")

        # Tab Resumo
        resumo_widget = QWidget()
        resumo_layout = QVBoxLayout(resumo_widget)
        resumo_layout.setContentsMargins(12, 12, 12, 12)
        resumo_layout.setSpacing(12)

        self._summary_cards_row = QHBoxLayout()
        self._summary_cards_row.setSpacing(10)
        resumo_layout.addLayout(self._summary_cards_row)

        self.summary_table = QTableWidget(0, 2)
        self.summary_table.setHorizontalHeaderLabels(["Métrica", "Valor"])
        _hh = self.summary_table.horizontalHeader()
        if _hh is not None:
            _hh.setStretchLastSection(True)
        _vh = self.summary_table.verticalHeader()
        if _vh is not None:
            _vh.setVisible(False)
        self.summary_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.summary_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        resumo_layout.addWidget(self.summary_table)

        self._no_data_label = QLabel("Nenhum relatório gerado ainda.")
        self._no_data_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._no_data_label.setStyleSheet("color: #8b949e; font-size: 13px; padding: 40px;")
        resumo_layout.addWidget(self._no_data_label)
        self.summary_table.setVisible(False)

        self.tabs.addTab(resumo_widget, "📊  Resumo")

        # Tab Configurações
        config_widget = QWidget()
        config_layout = QVBoxLayout(config_widget)
        config_layout.setContentsMargins(20, 16, 20, 16)
        config_layout.setSpacing(16)

        config_lbl = QLabel("CONFIGURAÇÕES DE CONSULTA")
        config_lbl.setObjectName("section")
        config_layout.addWidget(config_lbl)

        form = QFormLayout()
        form.setSpacing(12)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(5, 120)
        self.timeout_spin.setValue(10)
        self.timeout_spin.setSuffix("  seg")
        self.timeout_spin.setFixedWidth(120)
        self.timeout_spin.setToolTip("Tempo máximo de espera por requisição HTTP")
        form.addRow("Timeout:", self.timeout_spin)

        self.workers_spin = QSpinBox()
        self.workers_spin.setRange(1, 50)
        self.workers_spin.setValue(10)
        self.workers_spin.setSuffix("  threads")
        self.workers_spin.setFixedWidth(120)
        self.workers_spin.setToolTip("Número de consultas simultâneas")
        form.addRow("Máx. Workers:", self.workers_spin)

        config_layout.addLayout(form)
        config_layout.addStretch()
        self.tabs.addTab(config_widget, "⚙️  Configurações")

        root.addWidget(self.tabs, stretch=1)

    # ── Slots ────────────────────────────────────────────────────────

    def _browse_input(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo de entrada",
            "",
            "Planilhas (*.xlsx *.xls);;Texto (*.txt);;Todos (*)",
        )
        if path:
            self.input_edit.setText(path)
            out = str(Path(path).parent / "Relatorio_SITRAM.xlsx")
            self.output_edit.setText(out)

    def _browse_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar relatório como",
            "Relatorio_SITRAM.xlsx",
            "Planilha Excel (*.xlsx)",
        )
        if path:
            self.output_edit.setText(path)

    def _log(self, text: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        if "[AVISO]" in text:
            color = "#ff9800"
        elif any(w in text.upper() for w in ("ERRO", "FALHA", "INVÁLID", "CANCELADO")):
            color = "#f85149"
        elif any(w in text.lower() for w in ("sucesso", "gerado", "concluído")):
            color = "#3fb950"
        else:
            color = "#7ee787"
        self.log_area.append(
            f'<span style="color:#555;">[{ts}]</span> '
            f'<span style="color:{color};">{text}</span>'
        )
        self.log_area.moveCursor(QTextCursor.MoveOperation.End)

    def _update_progress(self, done: int, total: int) -> None:
        pct = int(done / total * 100) if total > 0 else 0
        self.progress_bar.setValue(pct)
        self.progress_label.setText(f"{done} / {total} chaves  ({pct}%)")
        self._status_bar.showMessage(f"Processando... {done}/{total} chaves ({pct}%)")

    def _tick_elapsed(self) -> None:
        if self._start_time is not None:
            delta = datetime.now() - self._start_time
            secs = int(delta.total_seconds())
            h, rem = divmod(secs, 3600)
            m, s = divmod(rem, 60)
            if h:
                self.elapsed_label.setText(f"⏱ {h:02d}:{m:02d}:{s:02d}")
            else:
                self.elapsed_label.setText(f"⏱ {m:02d}:{s:02d}")

    def _apply_config(self) -> None:
        try:
            import main as m
            m.TIMEOUT = self.timeout_spin.value()
            m.MAX_WORKERS = self.workers_spin.value()
        except Exception:
            pass

    def _start(self) -> None:
        input_file = self.input_edit.text().strip()
        output_file = self.output_edit.text().strip()

        if not input_file or not Path(input_file).exists():
            self._log("Arquivo de entrada não encontrado. Selecione um arquivo válido.")
            return
        if not output_file:
            self._log("Informe o caminho do arquivo de saída.")
            return

        self._apply_config()

        self._stop_event = threading.Event()
        self.run_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.open_btn.setVisible(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("")
        self.progress_label.setText("Iniciando...")
        self.log_area.clear()
        self._start_time = datetime.now()
        self._elapsed_timer.start()
        self._status_bar.showMessage("Iniciando processamento...")
        self.tabs.setCurrentIndex(0)

        self.worker = Worker(input_file, output_file, self._stop_event)
        self.worker.log_signal.connect(self._log)
        self.worker.progress_signal.connect(self._update_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.summary_signal.connect(self._update_summary)
        self.worker.start()

    def _cancel(self) -> None:
        if self.worker and self.worker.isRunning():
            self._stop_event.set()
            self.cancel_btn.setEnabled(False)
            self._log("[AVISO] Cancelamento solicitado. Aguardando tarefas em andamento...")
            self._status_bar.showMessage("Cancelando...")

    def _on_finished(self, success: bool, message: str) -> None:
        self._elapsed_timer.stop()
        self.run_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

        elapsed = ""
        if self._start_time:
            secs = int((datetime.now() - self._start_time).total_seconds())
            m, s = divmod(secs, 60)
            elapsed = f"{m:02d}:{s:02d}"

        if success:
            self.progress_bar.setValue(100)
            self.progress_bar.setStyleSheet(
                "QProgressBar::chunk { background-color: #238636; border-radius: 4px; }"
            )
            self.progress_label.setText(f"Concluído em {elapsed}!")
            self._log(f"Relatório gerado: {message}")
            self._output_path = message
            self.open_btn.setVisible(True)
            self._status_bar.showMessage(f"Concluído em {elapsed}  |  {message}")
        else:
            self.progress_bar.setStyleSheet(
                "QProgressBar::chunk { background-color: #f85149; border-radius: 4px; }"
            )
            self.progress_label.setText("Encerrado.")
            self._log(f"ERRO: {message}")
            self._status_bar.showMessage(f"Encerrado: {message}")

    def _update_summary(self, stats: dict) -> None:
        self._no_data_label.setVisible(False)
        self.summary_table.setVisible(True)

        # Limpa cards antigos
        while self._summary_cards_row.count():
            item = self._summary_cards_row.takeAt(0)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

        card_data = [
            ("Total\nChaves", str(stats["total_chaves"]), "#58a6ff"),
            ("Sucesso", str(stats["sucesso"]), "#3fb950"),
            ("Erros", str(stats["erros"]), "#f85149"),
            ("Lançamentos", str(stats["lancamentos"]), "#d2a8ff"),
            ("Itens", str(stats["itens"]), "#ffa657"),
        ]
        for label, value, color in card_data:
            card = QFrame()
            card.setObjectName("card")
            card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            cl = QVBoxLayout(card)
            cl.setContentsMargins(12, 10, 12, 10)
            cl.setSpacing(2)
            val_lbl = QLabel(value)
            val_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            val_lbl.setStyleSheet(f"color: {color}; font-size: 24px; font-weight: bold;")
            txt_lbl = QLabel(label)
            txt_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            txt_lbl.setStyleSheet("color: #8b949e; font-size: 11px;")
            cl.addWidget(val_lbl)
            cl.addWidget(txt_lbl)
            self._summary_cards_row.addWidget(card)

        # Tabela de métricas detalhadas
        rows = [
            ("Total de Chaves", str(stats["total_chaves"])),
            ("Notas com Sucesso", str(stats["sucesso"])),
            ("Notas com Erro", str(stats["erros"])),
            ("Total de Lançamentos", str(stats["lancamentos"])),
            ("Total de Itens", str(stats["itens"])),
            ("Valor Total das Notas", f"R$ {stats['valor_total']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")),
            ("Data de Geração", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ]
        self.summary_table.setRowCount(len(rows))
        for i, (key, val) in enumerate(rows):
            self.summary_table.setItem(i, 0, QTableWidgetItem(key))
            self.summary_table.setItem(i, 1, QTableWidgetItem(val))

        self.tabs.setCurrentIndex(1)

    def _open_output(self) -> None:
        path = self._output_path or self.output_edit.text()
        if Path(path).exists():
            os.startfile(path)


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
