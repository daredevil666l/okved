import sys
import requests
import pandas as pd
from itertools import cycle

from PyQt5.QtCore import Qt, QRect, QPropertyAnimation
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QLabel, QLineEdit, QCompleter, QFileDialog,
    QMessageBox, QPushButton, QGraphicsOpacityEffect
)

# Список API-ключей для ротации
API_KEYS = [
    '','MKXMk0GLaUN48g63','8zg1SLEpQPe7DCDH','K2TX2471T12E4qMP','8L1FghahT7pWWv8H','WgSDVbbjfakK4Pj2','Nx6JYT5tXlCDrT89','yIvRc7Z29tZ2pvmw','o7rA5pXvQvUBmkCT','UYtyJAYCFuCIitpu','ijo2nJdF2ZVMuyY8','bfcsAIOe0dS9bbeK','HZa0jsZSxUZI6x44','wros2k7udZCInXbh','DoBA8V5nSXidszTO','iBtUfMDQH3klPsdx','wWs0zpbtMAxAZoV1','TFOp6wBsqJxQD6sW','nvmRggGoc6av1XFn','gC7Vinvt74myvtB2','xmHSJECg4PnW35XR'

]
KEYS_CYCLE = cycle(API_KEYS)

SEARCH_URL = 'https://api.checko.ru/v2/search'
COMPANY_URL = 'https://api.checko.ru/v2/company'

OKVED_CODES = ['01.11','01.12','01.13','10.11','10.12','62.01','62.02','63.11','63.12','63.99']

def get_api_key():
    return next(KEYS_CYCLE)

def safe_request(url, params):
    """Пытаемся запросить, при ошибке меняем ключ и повторяем."""
    for _ in range(len(API_KEYS)):
        params['key'] = get_api_key()
        try:
            r = requests.get(url, params=params, timeout=5)
            data = r.json()
            if data.get('meta',{}).get('status') == 'ok':
                return data['data']
        except Exception:
            pass
    raise RuntimeError("Все API-ключи недоступны")

def fetch_organizations(okved):
    params = {'by':'okved','obj':'org','query':okved,'active':'true','limit':1000}
    orgs, page = [], 1
    while True:
        params['page'] = page
        data = safe_request(SEARCH_URL, params)
        recs = data.get('Записи',[])
        if not recs: break
        orgs.extend(recs)
        if page * params['limit'] >= data.get('СтрВсего',0): break
        page += 1
    return orgs

def fetch_contacts(ogrn):
    params = {'ogrn': ogrn}
    data = safe_request(COMPANY_URL, params)
    contacts = data.get('Контакты',{})
    return '; '.join(contacts.get('Тел',[])), '; '.join(contacts.get('Емэйл',[]))

class AnimatedButton(QPushButton):
    def __init__(self, text):
        super().__init__(text)
        self.anim = QPropertyAnimation(self, b"geometry")
        self.default_rect = None

    def enterEvent(self, e):
        if not self.default_rect:
            self.default_rect = self.geometry()
        end = self.default_rect.adjusted(-10, -5, 10, 5)
        self.anim.stop()
        self.anim.setDuration(200)
        self.anim.setStartValue(self.default_rect)
        self.anim.setEndValue(end)
        self.anim.start()
        super().enterEvent(e)

    def leaveEvent(self, e):
        if self.default_rect:
            self.anim.stop()
            self.anim.setDuration(200)
            self.anim.setStartValue(self.geometry())
            self.anim.setEndValue(self.default_rect)
            self.anim.start()
        super().leaveEvent(e)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Notice")
        self.setMinimumSize(400, 250)

        # Заголовок программы
        self.title_label = QLabel("Notice", self)
        font = QFont('Segoe UI', 32, QFont.Bold)
        self.title_label.setFont(font)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setGeometry(0, 10, 150, 25)

        # Анимация заголовка
        self.title_anim = QPropertyAnimation(self.title_label, b"geometry")
        self.title_anim.setDuration(800)
        self.title_anim.setStartValue(QRect(-800, 20, 800, 50))
        self.title_anim.setEndValue(QRect(0, 10, 150, 25))
        self.title_anim.start()

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(50, 100, 50, 50)
        layout.setSpacing(20)

        layout.addWidget(QLabel("Введите код ОКВЭД:"))
        self.okved_input = QLineEdit()
        completer = QCompleter(OKVED_CODES)
        completer.setCaseSensitivity(False)
        self.okved_input.setCompleter(completer)
        self.okved_input.setFixedHeight(30)
        layout.addWidget(self.okved_input)

        self.search_btn = AnimatedButton("Поиск и сохранение")
        self.search_btn.setFixedHeight(40)
        self.search_btn.clicked.connect(self.on_search)
        layout.addWidget(self.search_btn)

        self.setCentralWidget(central)

    def on_search(self):
        okved = self.okved_input.text().strip()
        if not okved:
            QMessageBox.warning(self, "Ошибка", "Код ОКВЭД не введён")
            return
        try:
            orgs = fetch_organizations(okved)
        except RuntimeError as e:
            QMessageBox.critical(self, "Ошибка API", str(e))
            return

        if not orgs:
            QMessageBox.information(self, "Результат", "Ничего не найдено")
            return

        rows = []
        for rec in orgs:
            ogrn, name, inn = rec['ОГРН'], rec['НаимПолн'], rec['ИНН']
            phones, emails = fetch_contacts(ogrn)
            rows.append({'Название': name, 'ИНН': inn, 'Телефоны': phones, 'E-mail': emails})

        path, _ = QFileDialog.getSaveFileName(self, "Сохранить", f"orgs_{okved}.xlsx", "Excel Files (*.xlsx)")
        if path:
            pd.DataFrame(rows).to_excel(path, index=False)
            QMessageBox.information(self, "Готово", f"Сохранено в {path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
