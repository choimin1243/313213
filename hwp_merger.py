import sys
import os
import zipfile
import shutil
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem, QLabel, QFileDialog,
    QMessageBox, QProgressBar, QAbstractItemView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent


class MergeWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, file_paths, output_path):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path

    def run(self):
        try:
            ext = Path(self.file_paths[0]).suffix.lower()
            if ext == '.hwpx':
                self._merge_hwpx()
            elif ext == '.hwp':
                self._merge_hwp()
        except Exception as e:
            self.error.emit(str(e))

    def _merge_hwpx(self):
        """HWPX(ZIP 기반) 파일 합치기"""
        import xml.etree.ElementTree as ET

        self.progress.emit(0, "첫 번째 파일 읽는 중...")
        base_path = self.file_paths[0]

        # 임시 작업 디렉토리
        tmp_dir = Path(self.output_path).parent / "_hwpx_tmp"
        if tmp_dir.exists():
            shutil.rmtree(tmp_dir)
        tmp_dir.mkdir()

        try:
            # 첫 번째 파일을 기준으로 압축 해제
            base_dir = tmp_dir / "base"
            with zipfile.ZipFile(base_path, 'r') as z:
                z.extractall(base_dir)

            # 기준 파일의 BodyText 섹션 파일 목록 수집
            body_dir = base_dir / "Contents"
            section_files = sorted(body_dir.glob("Section*.xml")) if body_dir.exists() else []
            section_count = len(section_files)

            for i, fp in enumerate(self.file_paths[1:], 1):
                self.progress.emit(
                    int(i / len(self.file_paths) * 80),
                    f"{Path(fp).name} 합치는 중..."
                )
                add_dir = tmp_dir / f"add_{i}"
                with zipfile.ZipFile(fp, 'r') as z:
                    z.extractall(add_dir)

                add_body = add_dir / "Contents"
                add_sections = sorted(add_body.glob("Section*.xml")) if add_body.exists() else []

                for sec in add_sections:
                    new_name = f"Section{section_count}.xml"
                    shutil.copy(sec, body_dir / new_name)
                    section_count += 1

            # contents.hpf (목차) 업데이트
            hpf_path = base_dir / "Contents" / "content.hpf"
            if not hpf_path.exists():
                hpf_path = base_dir / "content.hpf"

            if hpf_path.exists():
                tree = ET.parse(hpf_path)
                root = tree.getroot()
                ns = {'opf': 'http://www.idpf.org/2007/opf'}
                manifest = root.find('.//manifest', ns) or root.find('.//manifest')
                spine = root.find('.//spine', ns) or root.find('.//spine')

                if manifest is not None and spine is not None:
                    # 새로 추가된 섹션 등록
                    existing_items = {item.get('href') for item in manifest}
                    for idx in range(section_count):
                        href = f"Section{idx}.xml"
                        if href not in existing_items:
                            item_id = f"section{idx}"
                            ET.SubElement(manifest, 'item', {
                                'id': item_id,
                                'href': href,
                                'media-type': 'application/xml'
                            })
                            ET.SubElement(spine, 'itemref', {'idref': item_id})
                    tree.write(hpf_path, encoding='utf-8', xml_declaration=True)

            # 결과물 압축
            self.progress.emit(90, "파일 저장 중...")
            output = self.output_path
            if not output.endswith('.hwpx'):
                output += '.hwpx'

            with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
                for f in base_dir.rglob('*'):
                    if f.is_file():
                        zout.write(f, f.relative_to(base_dir))

            self.progress.emit(100, "완료!")
            self.finished.emit(output)

        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def _merge_hwp(self):
        """HWP(바이너리) 파일 합치기 - hwp5 라이브러리 사용"""
        try:
            import hwp5
            from hwp5.xmlmodel import Hwp5File
        except ImportError:
            self.error.emit(
                "HWP 바이너리 형식 처리를 위해 pyhwp 라이브러리가 필요합니다.\n"
                "pip install pyhwp 를 실행해주세요.\n\n"
                "※ HWP 형식은 바이너리 구조상 완전한 합치기가 제한적입니다.\n"
                "가능하면 HWPX 형식(.hwpx)으로 저장 후 사용하세요."
            )
            return

        # pyhwp 기반 처리 (섹션 단위 병합)
        self.progress.emit(50, "HWP 파일 처리 중 (제한적 지원)...")
        self.error.emit(
            "HWP 바이너리 형식은 완전한 자동 합치기가 어렵습니다.\n"
            "한컴오피스에서 파일을 HWPX 형식으로 저장 후 다시 시도해주세요."
        )


class DropListWidget(QListWidget):
    """드래그&드롭 지원 리스트"""
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith(('.hwp', '.hwpx')):
                    self._add_file(path)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def _add_file(self, path):
        # 중복 방지
        for i in range(self.count()):
            if self.item(i).data(Qt.UserRole) == path:
                return
        item = QListWidgetItem(f"📄 {Path(path).name}")
        item.setData(Qt.UserRole, path)
        item.setToolTip(path)
        self.addItem(item)


class HwpMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HWP / HWPX 파일 합치기")
        self.setMinimumSize(600, 500)
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        # 안내 레이블
        info = QLabel("📂 HWP / HWPX 파일을 추가하고 순서를 정렬한 뒤 합치기를 실행하세요.\n"
                      "파일을 드래그&드롭하거나 아래 버튼으로 추가할 수 있습니다.")
        info.setWordWrap(True)
        layout.addWidget(info)

        # 파일 목록
        self.list_widget = DropListWidget()
        layout.addWidget(self.list_widget)

        # 버튼 행 1 - 파일 관리
        btn_row1 = QHBoxLayout()
        self.btn_add = QPushButton("➕ 파일 추가")
        self.btn_up = QPushButton("⬆ 위로")
        self.btn_down = QPushButton("⬇ 아래로")
        self.btn_remove = QPushButton("🗑 선택 삭제")
        self.btn_clear = QPushButton("✖ 전체 삭제")
        for btn in [self.btn_add, self.btn_up, self.btn_down, self.btn_remove, self.btn_clear]:
            btn_row1.addWidget(btn)
        layout.addLayout(btn_row1)

        # 버튼 행 2 - 실행
        btn_row2 = QHBoxLayout()
        self.btn_merge = QPushButton("🔗 합치기 실행")
        self.btn_merge.setFixedHeight(40)
        self.btn_merge.setStyleSheet("font-size: 14px; font-weight: bold; background-color: #4CAF50; color: white;")
        btn_row2.addWidget(self.btn_merge)
        layout.addLayout(btn_row2)

        # 진행바
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        # 시그널 연결
        self.btn_add.clicked.connect(self.add_files)
        self.btn_up.clicked.connect(self.move_up)
        self.btn_down.clicked.connect(self.move_down)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.clear_all)
        self.btn_merge.clicked.connect(self.run_merge)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "HWP/HWPX 파일 선택", "",
            "한글 파일 (*.hwp *.hwpx);;HWP (*.hwp);;HWPX (*.hwpx)"
        )
        for f in files:
            self.list_widget._add_file(f)

    def move_up(self):
        row = self.list_widget.currentRow()
        if row > 0:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row - 1, item)
            self.list_widget.setCurrentRow(row - 1)

    def move_down(self):
        row = self.list_widget.currentRow()
        if row < self.list_widget.count() - 1:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row + 1, item)
            self.list_widget.setCurrentRow(row + 1)

    def remove_selected(self):
        for item in self.list_widget.selectedItems():
            self.list_widget.takeItem(self.list_widget.row(item))

    def clear_all(self):
        self.list_widget.clear()

    def get_file_paths(self):
        return [
            self.list_widget.item(i).data(Qt.UserRole)
            for i in range(self.list_widget.count())
        ]

    def run_merge(self):
        paths = self.get_file_paths()
        if len(paths) < 2:
            QMessageBox.warning(self, "경고", "합칠 파일을 2개 이상 추가해주세요.")
            return

        # 확장자 일관성 체크
        exts = set(Path(p).suffix.lower() for p in paths)
        if len(exts) > 1:
            QMessageBox.warning(self, "경고", "HWP와 HWPX 파일을 혼합할 수 없습니다.\n같은 형식의 파일만 선택해주세요.")
            return

        ext = exts.pop()
        output, _ = QFileDialog.getSaveFileName(
            self, "저장 위치 선택", f"merged{ext}",
            f"한글 파일 (*{ext})"
        )
        if not output:
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.btn_merge.setEnabled(False)

        self.worker = MergeWorker(paths, output)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, value, msg):
        self.progress_bar.setValue(value)
        self.status_label.setText(msg)

    def on_finished(self, output_path):
        self.progress_bar.setValue(100)
        self.btn_merge.setEnabled(True)
        QMessageBox.information(self, "완료", f"파일이 저장되었습니다:\n{output_path}")
        self.status_label.setText(f"✅ 저장 완료: {output_path}")

    def on_error(self, msg):
        self.progress_bar.setVisible(False)
        self.btn_merge.setEnabled(True)
        QMessageBox.critical(self, "오류", msg)
        self.status_label.setText("❌ 오류 발생")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = HwpMergerApp()
    window.show()
    sys.exit(app.exec_())
