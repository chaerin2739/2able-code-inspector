import sys
import pandas as pd
import re
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QTreeWidgetItem, QTableWidgetItem, QMenu
from PyQt5.QtCore import Qt
import os

class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        ui_file_path = os.path.join(os.path.dirname(__file__), 'main.ui')
        uic.loadUi(ui_file_path, self)  # UI 파일 로드
        # 변수 저장
        self.excel_data = None  # 엑셀 데이터를 저장할 변수
        self.python_lines = None  # 파이썬 파일의 줄들을 저장할 변수

        # 버튼 클릭 시 연결할 함수 설정
        self.pushButton_2.clicked.connect(self.load_excel_file)  # 엑셀 파일 불러오기
        self.pushButton.clicked.connect(self.load_python_file)   # 파이썬 파일 불러오기
        self.pushButton_3.clicked.connect(self.run_check_script)  # 변경 기능 수행

        # treeWidget 항목 더블 클릭 시 파이썬 코드에서 해당 항목으로 이동하는 기능 연결
        self.treeWidget.itemDoubleClicked.connect(self.open_edit_tree_item)
        self.treeWidget.itemChanged.connect(self.close_edit_tree_item)

        # 오른쪽 클릭 시 컨텍스트 메뉴 연결
        self.treeWidget.setContextMenuPolicy(3)  # Qt.CustomContextMenu
        self.treeWidget.customContextMenuRequested.connect(self.show_context_menu)

        # 변경하기 버튼 연결
        self.pushButton_4.clicked.connect(self.update_tree_widget)  # 변경하기 버튼 클릭 연결

    def load_excel_file(self):
        try:
            # 엑셀 파일 선택 및 불러오기
            options = QFileDialog.Options()
            file_name, _ = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', '', 'CSV Files (*.csv);;All Files (*)', options=options)
            if file_name:
                # CSV 파일 읽기 (첫 줄 건너뛰기)
                self.excel_data = pd.read_csv(file_name, header=None, names=['제목', '내용', '변경 제목', '변경 내용'])
                self.excel_data = self.excel_data.iloc[1:].reset_index(drop=True)
                # treeWidget 초기화
                self.treeWidget.clear()

                # CSV 파일의 각 행을 treeWidget에 한 줄씩 추가
                for idx, row in self.excel_data.iterrows():
                    item = QTreeWidgetItem([row['제목'], row['내용'], row['변경 제목'], row['변경 내용']])
                    item.setFlags(item.flags() | 2)  # 아이템 편집 가능 설정
                    self.treeWidget.addTopLevelItem(item)

                QMessageBox.information(self, '성공', f'{file_name} 파일을 성공적으로 불러왔습니다.')
        except Exception as e:
            QMessageBox.warning(self, '오류', f'엑셀 파일을 불러오는 중 오류가 발생했습니다: {e}')

    def load_python_file(self):
        try:
            # 파이썬 파일 선택 및 불러오기
            options = QFileDialog.Options()
            file_name, _ = QFileDialog.getOpenFileName(self, '파이썬 파일 선택', '', 'Python Files (*.py);;All Files (*)', options=options)
            if file_name:
                # 파이썬 파일 읽기
                with open(file_name, 'r', encoding='utf-8') as file:
                    self.python_lines = file.readlines()

                # tableWidget 초기화
                self.tableWidget.setRowCount(0)  # 기존 내용 초기화

                # 파이썬 파일의 각 줄을 tableWidget에 한 줄씩 추가
                for idx, line in enumerate(self.python_lines):
                    self.tableWidget.insertRow(idx)
                    item = QTableWidgetItem(line.strip())  # 줄바꿈 문자를 제거한 후 추가
                    self.tableWidget.setItem(idx, 0, item)
                    self.tableWidget.setColumnWidth(0, max(self.tableWidget.columnWidth(0), len(line.strip()) * 7))  # 셀 너비 조정

                QMessageBox.information(self, '성공', f'{file_name} 파일을 성공적으로 불러왔습니다.')
        except Exception as e:
            QMessageBox.warning(self, '오류', f'파이썬 파일을 불러오는 중 오류가 발생했습니다: {e}')

    def highlight_python_code(self, item):
        try:
            # 더블 클릭된 엑셀 항목에서 제목과 내용을 가져옴
            clicked_title = item.text(0)
            clicked_content = item.text(1)

            # 파이썬 파일의 줄을 확인하여 해당 제목과 내용이 있는 곳을 찾음
            for idx, line in enumerate(self.python_lines):
                if clicked_title in line and clicked_content in line:
                    # 해당 줄을 tableWidget에서 선택
                    self.tableWidget.selectRow(idx)
                    self.tableWidget.scrollToItem(self.tableWidget.item(idx, 0))  # 해당 줄로 스크롤 이동
                    break
        except Exception as e:
            QMessageBox.warning(self, '오류', f'해당 항목을 찾는 중 오류가 발생했습니다: {e}')

    def run_check_script(self):
        try:
            if self.excel_data is None or self.python_lines is None:
                QMessageBox.warning(self, '오류', '엑셀 파일과 파이썬 파일을 먼저 불러오세요.')
                return

            # 하이라이트된 행 선택 확인
            selected_rows = list(set(index.row() for index in self.tableWidget.selectedIndexes()))
            if not selected_rows:
                QMessageBox.warning(self, '오류', '먼저 하이라이트할 파이썬 부분을 선택하세요.')
                return

            # 선택된 행의 내용을 가져와서 변경
            for row in selected_rows:
                original_line = self.python_lines[row]
                original_title_match = re.search(r'imsg\.show_modal\(\s*[\'"](.+?)[\'"]\s*,\s*(f?[\'"](.+?)[\'"])\s*\)', original_line)

                if original_title_match:
                    # 기존 제목과 내용을 가져옴
                    current_title = original_title_match.group(1)
                    current_content = original_title_match.group(2)

                    # CSV에서 제목과 내용이 일치하는 행을 찾습니다.
                    row_data = self.excel_data[(self.excel_data['제목'] == current_title) & 
                                                (self.excel_data['내용'].str.strip() == current_content.strip('f"').strip('\''))]

                    if not row_data.empty:
                        new_title = row_data['변경 제목'].values[0]
                        new_content = row_data['변경 내용'].values[0]

                        # 변경된 코드 생성 (기존 들여쓰기 유지)
                        indent = len(original_line) - len(original_line.lstrip())
                        new_line = ' ' * indent + f'imsg.show_modal("{new_title}", "{new_content}")\n'
                        self.python_lines[row] = new_line

            # 변경된 코드를 tableWidget에 업데이트
            self.tableWidget.setRowCount(0)  # tableWidget 초기화
            for idx, line in enumerate(self.python_lines):
                self.tableWidget.insertRow(idx)
                self.tableWidget.setItem(idx, 0, QTableWidgetItem(line.strip()))

            # 새로운 파일에 저장
            with open('runrun.py', 'w', encoding='utf-8') as new_file:
                new_file.writelines(self.python_lines)

            QMessageBox.information(self, '완료', '변경 작업이 완료되었습니다.')
        except Exception as e:
            QMessageBox.warning(self, '오류', f'변경 작업 중 오류가 발생했습니다: {e}')

    def close_edit_tree_item(self, item, column):
        # 항목이 변경되면 편집 모드를 닫음
        self.treeWidget.closePersistentEditor(item, column)

    def open_edit_tree_item(self, item, column):
        # 항목 더블클릭 시 편집 모드를 열음
        self.treeWidget.openPersistentEditor(item, column)

        # 커서가 다른 곳으로 이동했을 때 닫히도록 설정
        self.treeWidget.itemChanged.emit(item, column)  # 강제로 itemChanged 이벤트 발생시킴

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
    # 모든 항목의 편집기를 닫음
        for index in range(self.treeWidget.topLevelItemCount()):
            item = self.treeWidget.topLevelItem(index)
            self.treeWidget.closePersistentEditor(item)

    def update_tree_widget(self):
        try:
            selected_item = self.treeWidget.currentItem()
            if selected_item:
                # 현재 항목에서 제목과 내용을 수정
                current_title = selected_item.text(0)
                current_content = selected_item.text(1)

                # 변경된 제목과 내용을 가져오기
                new_title = selected_item.text(2)
                new_content = selected_item.text(3)

                # 엑셀 데이터 업데이트
                row_index = self.excel_data.index[(self.excel_data['제목'] == current_title) & 
                                                   (self.excel_data['내용'] == current_content)].tolist()
                if row_index:
                    self.excel_data.at[row_index[0], '변경 제목'] = new_title
                    self.excel_data.at[row_index[0], '변경 내용'] = new_content

                    # treeWidget 업데이트
                    selected_item.setText(2, new_title)
                    selected_item.setText(3, new_content)

                    # 업데이트된 엑셀 파일 저장
                    self.excel_data.to_csv('updated_excel.csv', index=False, header=False)
                    QMessageBox.information(self, '완료', '엑셀 파일이 업데이트되었습니다.')
                else:
                    QMessageBox.warning(self, '오류', '해당 항목을 엑셀 데이터에서 찾을 수 없습니다.')
        except Exception as e:
            QMessageBox.warning(self, '오류', f'변경 작업 중 오류가 발생했습니다: {e}')

    def show_context_menu(self, pos):
        context_menu = QMenu(self)
        open_python_action = context_menu.addAction("파이썬 코드로 이동")
        action = context_menu.exec_(self.treeWidget.viewport().mapToGlobal(pos))

        if action == open_python_action:
            selected_item = self.treeWidget.currentItem()
            if selected_item:
                self.highlight_python_code(selected_item)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec_())
