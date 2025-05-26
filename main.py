import sys
import os
import platform
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QFileDialog, QMessageBox, QLabel,
    QSplitter, QTextEdit, QStatusBar
)
from PyQt5.QtCore import QUrl, pyqtSlot, QDir, QStandardPaths, Qt, QSize, QTime
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineProfile, QWebEngineDownloadItem
from PyQt5.QtGui import QIcon, QDesktopServices

# JS-код с убранным ограничением на выгрузку
JS_CODE = r"""
function exportToCsv(e,t){
    for(var n="",o=0;o<t.length;o++)n+=function(e){for(var t="",n=0;n<e.length;n++){var o=null===e[n]||void 0===e[n]?"":e[n].toString(),o=(o=e[n]instanceof Date?e[n].toLocaleString():o).replace(/"/g,'""');0<n&&(t+=","),t+=o=0<=o.search(/("|,|\n)/g)?'"'+o+'"':o}return t+"\n"}(t[o]);var r=new Blob([n],{type:"text/csv;charset=utf-8;"}),i=document.createElement("a");void 0!==i.download&&(r=URL.createObjectURL(r),i.setAttribute("href",r),i.setAttribute("download",e),document.body.appendChild(i),i.click(),document.body.removeChild(i))
}

function buildCTABtn(){
    if (document.getElementById('fb-group-scraper-button')) {
        return document.getElementById('fb-group-scraper-container');
    }

    var e=document.createElement("div");
    e.id = 'fb-group-scraper-container';
    e.setAttribute("style",["position: fixed;","top: 0;","left: 0;","z-index: 10000;","width: 100%;","height: 100%;","pointer-events: none;"].join(""));

    var t = document.createElement("div");
    t.id = 'fb-group-scraper-button';
    t.setAttribute("style",["position: absolute;","bottom: 30px;","right: 130px;","color: white;","min-width: 150px;","background: #1877F2;","border-radius: 6px;","padding: 0px 12px;","cursor: pointer;","font-weight:600;","font-size:15px;","display: inline-flex;","pointer-events: auto;","height: 36px;","align-items: center;","justify-content: center;","box-shadow: 0 2px 8px rgba(0,0,0,0.2);","z-index: 10001;"].join(""));

    var n = document.createTextNode("Download ");
    var o = document.createElement("span");
    o.setAttribute("id","fb-group-scraper-number-tracker");
    o.textContent = "0";
    var r = document.createTextNode(" members");

    t.appendChild(n);
    t.appendChild(o);
    t.appendChild(r);

    t.addEventListener("click",function(){
        var e = (new Date).toISOString().replace(/:/g, '-');
        console.log("Exporting members:", window.members_list.length);
        exportToCsv("groupMemberExport-"+e+".csv", window.members_list);
    });

    var excelBtn = document.createElement("div");
    excelBtn.id = 'fb-group-scraper-excel-button';
    excelBtn.setAttribute("style",["position: absolute;","bottom: 30px;","right: 300px;","color: white;","min-width: 150px;","background: #217346;","border-radius: 6px;","padding: 0px 12px;","cursor: pointer;","font-weight:600;","font-size:15px;","display: inline-flex;","pointer-events: auto;","height: 36px;","align-items: center;","justify-content: center;","box-shadow: 0 2px 8px rgba(0,0,0,0.2);","z-index: 10001;"].join(""));

    var excelText = document.createTextNode("Excel Export ");
    var excelNumSpan = document.createElement("span");
    excelNumSpan.setAttribute("id","fb-group-scraper-excel-number-tracker");
    excelNumSpan.textContent = "0";
    var excelSuffix = document.createTextNode(" members");

    excelBtn.appendChild(excelText);
    excelBtn.appendChild(excelNumSpan);
    excelBtn.appendChild(excelSuffix);

    excelBtn.addEventListener("click",function(){
        var e = (new Date).toISOString().replace(/:/g, '-');
        window.pyExcelExport = true;
        console.log("Exporting to Excel:", window.members_list.length);
        exportToCsv("groupMemberExport-"+e+".csv", window.members_list);
    });

    e.appendChild(t);
    e.appendChild(excelBtn);
    document.body.appendChild(e);
    return e;
}

function processResponse(e){
    var t, n, o;

    if(null!==(t=null==e?void 0:e.data)&&void 0!==t&&t.group)
        o=e.data.group;
    else{
        if("Group"!==(null===(t=null===(t=null==e?void 0:e.data)||void 0===t?void 0:t.node)||void 0===t?void 0:t.__typename))
            return;
        o=e.data.node;
    }

    if(null!==(t=null==o?void 0:o.new_members)&&void 0!==t&&t.edges)
        n=o.new_members.edges;
    else if(null!==(e=null==o?void 0:o.new_forum_members)&&void 0!==e&&e.edges)
        n=o.new_forum_members.edges;
    else{
        if(null===(t=null==o?void 0:o.search_results)||void 0===t||!t.edges)
            return;
        n=o.search_results.edges;
    }

    var e=n.map(function(e){
        var t=e.node,
            n=t.id,
            o=t.name,
            r=t.bio_text,
            i=t.url,
            s=t.profile_picture,
            t=t.__isProfile,
            d=(null===(d=null==e?void 0:e.join_status_text)||void 0===d?void 0:d.text)||
              (null===(d=null===(d=null==e?void 0:e.membership)||void 0===d?void 0:d.join_status_text)||void 0===d?void 0:d.text),
            e=null===(e=e.node.group_membership)||void 0===e?void 0:e.associated_group.id;

        return[n,o,i,(null==r?void 0:r.text)||"",(null==s?void 0:s.uri)||"",e,d||"",t]
    });

    ((t=window.members_list).push.apply(t,e));
    var o=document.getElementById("fb-group-scraper-number-tracker");
    var excelCounter = document.getElementById("fb-group-scraper-excel-number-tracker");

    if(o) {
        var count = window.members_list.length - 1;
        o.textContent = count.toString();
        excelCounter.textContent = count.toString();
        console.log("Total members collected:", count);
    }
}

function parseResponse(e){
    var n=[];
    try{
        n.push(JSON.parse(e));
    }catch(t){
        var o=e.split("\n");
        if(o.length<=1) {
            console.error("Failed to parse API response", t);
            return;
        }

        for(var r=0;r<o.length;r++){
            var i=o[r];
            if (i.trim() === '') continue;
            try{
                n.push(JSON.parse(i));
            }catch(e){
                console.error("Failed to parse part of API response", e);
            }
        }
    }

    for(var t=0;t<n.length;t++)
        processResponse(n[t]);
}

function main(){
    console.log("FB Group Exporter started!");
    window.members_list = window.members_list || [["Profile Id","Full Name","ProfileLink","Bio","Image Src","Group Id","Group Joining Text","Profile Type"]];
    window.pyExcelExport = false;

    var btnContainer = buildCTABtn();

    var originalXHRSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function(){
        this.addEventListener("readystatechange", function(){
            if(this.responseURL.includes("/api/graphql/") && this.readyState === 4) {
                parseResponse(this.responseText);
            }
        }, false);
        originalXHRSend.apply(this, arguments);
    };

    console.log("FB Group Exporter initialized, scraping started");
}

main();
"""

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FB Group Exporter")
        self.resize(1200, 800)

        self.is_mac = platform.system() == 'Darwin'
        self.download_dir = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)

        self.url_input = QLineEdit("https://www.facebook.com/groups/YOUR_GROUP_ID/members/")
        self.url_input.setPlaceholderText("Введите URL группы Facebook")
        self.load_btn = QPushButton("Загрузить группу")
        self.inject_btn = QPushButton("Внедрить скрипт")
        self.inject_btn.setEnabled(False)
        self.choose_dir_btn = QPushButton("Выбрать папку")
        self.contact_btn = QPushButton("Задать вопрос")

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(80)
        self.log_area.setPlaceholderText("История операций будет отображаться здесь")

        self.status_bar = QStatusBar()
        self.status_bar.setSizeGripEnabled(False)
        self.status_bar.showMessage(f"Папка сохранения: {self.download_dir}")

        self.view = QWebEngineView()
        self.view.loadFinished.connect(self.on_page_loaded)
        self.view.setMinimumHeight(500)

        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.addWidget(self.view)
        self.splitter.addWidget(self.log_area)
        self.splitter.setStretchFactor(0, 4)
        self.splitter.setStretchFactor(1, 1)

        topbar = QHBoxLayout()
        topbar.addWidget(self.url_input, 3)
        topbar.addWidget(self.load_btn)
        topbar.addWidget(self.inject_btn)
        topbar.addWidget(self.choose_dir_btn)
        topbar.addWidget(self.contact_btn)

        layout = QVBoxLayout(self)
        layout.addLayout(topbar)
        layout.addWidget(self.splitter, 1)
        layout.addWidget(self.status_bar)

        self.load_btn.clicked.connect(self.load_page)
        self.inject_btn.clicked.connect(self.inject_js)
        self.choose_dir_btn.clicked.connect(self.choose_download_dir)
        self.contact_btn.clicked.connect(self.open_contact)

        self.setup_download_handler()

        self.setStyleSheet("""
            QPushButton {
                background-color: #1877F2;
                color: white;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #166fe5;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
            QLineEdit {
                padding: 6px;
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            QTextEdit {
                border: 1px solid #ccc;
                border-radius: 4px;
            }
        """)

        self.log("Программа запущена. Готова к работе.")

    def log(self, message):
        self.log_area.append(f"[{QTime.currentTime().toString('HH:mm:ss')}] {message}")
        self.log_area.ensureCursorVisible()
        self.status_bar.showMessage(message, 5000)

    def setup_download_handler(self):
        profile = QWebEngineProfile.defaultProfile()
        profile.downloadRequested.connect(self.handle_download)

        if self.is_mac:
            profile.setHttpUserAgent("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        try:
            if hasattr(profile, 'setDownloadPath'):
                profile.setDownloadPath(self.download_dir)
        except Exception as e:
            self.log(f"Примечание: Не удалось установить путь загрузки: {str(e)}")

    @pyqtSlot()
    def open_contact(self):
        QDesktopServices.openUrl(QUrl("https://t.me/EcommerceGr"))
        self.log("Открыта страница для связи с разработчиком")

    @pyqtSlot()
    def load_page(self):
        url = self.url_input.text().strip()
        if not url.startswith(("http://", "https://")):
            QMessageBox.warning(self, "Некорректный URL", "Введите корректный URL группы Facebook.")
            return

        self.log(f"Загрузка страницы: {url}")
        self.view.load(QUrl(url))

    @pyqtSlot(bool)
    def on_page_loaded(self, success):
        if success:
            self.inject_btn.setEnabled(True)
            self.log("Страница успешно загружена")
            if self.is_mac:
                self.view.page().runJavaScript("document.body.style.backgroundColor = 'white';")
        else:
            self.log("Не удалось загрузить страницу")

    @pyqtSlot()
    def inject_js(self):
        current_url = self.view.url().toString()
        if 'facebook.com/groups/' not in current_url:
            QMessageBox.warning(self, "Неверная страница",
                                "Убедитесь, что вы находитесь на странице участников группы Facebook.")
            return

        self.log("Внедрение скрипта...")
        self.view.page().runJavaScript(JS_CODE, self.on_js_injected)

    def on_js_injected(self, result):
        self.log(
            "Скрипт успешно внедрен! Найдите синюю кнопку 'Download X members' и зеленую 'Excel Export X members' в нижней части экрана")
        QMessageBox.information(self, "Скрипт внедрен",
                                "Скрипт успешно внедрен! Теперь:\n"
                                "1. Прокрутите страницу вниз для загрузки участников\n"
                                "2. Найдите синюю кнопку 'Download X members' или зеленую 'Excel Export X members'\n"
                                "3. Нажмите на нее для скачивания файла")

    @pyqtSlot()
    def choose_download_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "Выберите папку для сохранения",
            self.download_dir,
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )

        if dir_path:
            self.download_dir = dir_path
            self.log(f"Установлена новая папка сохранения: {self.download_dir}")
            try:
                profile = QWebEngineProfile.defaultProfile()
                if hasattr(profile, 'setDownloadPath'):
                    profile.setDownloadPath(self.download_dir)
            except Exception as e:
                self.log(f"Примечание: {str(e)}")

    @pyqtSlot(QWebEngineDownloadItem)
    def handle_download(self, download):
        if not self.download_dir or not os.path.isdir(self.download_dir):
            QMessageBox.warning(self, "Папка не найдена",
                                "Выбранная папка не существует. Пожалуйста, выберите другую.")
            return

        filename = download.downloadFileName()
        file_path = os.path.join(self.download_dir, filename)

        if self.is_mac:
            file_path = file_path.replace('\\', '/')

        download.setPath(file_path)
        download.finished.connect(lambda: self.on_download_finished(file_path))
        download.downloadProgress.connect(self.on_download_progress)
        download.accept()
        self.log(f"Начато скачивание файла: {filename}")

    def on_download_finished(self, file_path):
        if os.path.exists(file_path):
            self.log(f"Файл успешно сохранен: {file_path}")
            self.view.page().runJavaScript("window.pyExcelExport", self.convert_to_excel_if_needed(file_path))
            self.status_bar.showMessage(f"Файл успешно сохранен!", 10000)
        else:
            self.log(f"Ошибка при сохранении файла: {file_path}")
            QMessageBox.warning(self, "Ошибка скачивания",
                                "Файл не был сохранен. Проверьте разрешения папки.")

    def convert_to_excel_if_needed(self, csv_path):
        def convert_callback(need_excel):
            if need_excel:
                try:
                    self.view.page().runJavaScript("window.pyExcelExport = false;")
                    if not os.path.exists(csv_path):
                        self.log("Ошибка: CSV файл не найден для конвертации в Excel")
                        return

                    excel_path = csv_path.replace('.csv', '.xlsx')

                    import csv
                    try:
                        import xlsxwriter
                    except ImportError:
                        self.log("Для экспорта в Excel требуется библиотека xlsxwriter. Устанавливаем...")
                        try:
                            import subprocess
                            pip_cmd = [sys.executable, "-m", "pip", "install", "xlsxwriter", "--user"]
                            subprocess.check_call(pip_cmd)
                            import xlsxwriter
                            self.log("Библиотека xlsxwriter успешно установлена")
                        except Exception as e:
                            self.log(f"Ошибка при установке xlsxwriter: {str(e)}")
                            QMessageBox.warning(self, "Ошибка",
                                                "Невозможно установить библиотеку xlsxwriter.\n"
                                                "Пожалуйста, установите её вручную: pip3 install xlsxwriter --user")
                            return

                    rows = []
                    try:
                        with open(csv_path, 'r', encoding='utf-8') as csv_file:
                            csv_reader = csv.reader(csv_file)
                            for row in csv_reader:
                                rows.append(row)
                    except UnicodeDecodeError:
                        with open(csv_path, 'r', encoding='latin-1') as csv_file:
                            csv_reader = csv.reader(csv_file)
                            for row in csv_reader:
                                rows.append(row)

                    workbook = xlsxwriter.Workbook(excel_path)
                    worksheet = workbook.add_worksheet('Участники группы')

                    header_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#D9EAD3',
                        'border': 1
                    })

                    for row_idx, row in enumerate(rows):
                        for col_idx, value in enumerate(row):
                            if row_idx == 0:
                                worksheet.write(row_idx, col_idx, value, header_format)
                            else:
                                worksheet.write(row_idx, col_idx, value)

                    for col_idx in range(len(rows[0]) if rows else 0):
                        max_width = 0
                        for row_idx in range(len(rows)):
                            cell_value = rows[row_idx][col_idx] if col_idx < len(rows[row_idx]) else ""
                            if cell_value:
                                width = min(len(str(cell_value)), 100)
                                max_width = max(max_width, width)
                        worksheet.set_column(col_idx, col_idx, max_width + 1)

                    workbook.close()
                    self.log(f"Создан Excel файл: {excel_path}")
                    QMessageBox.information(self, "Экспорт в Excel",
                                            f"Excel файл успешно создан:\n{excel_path}")

                except Exception as e:
                    self.log(f"Ошибка при конвертации в Excel: {str(e)}")
                    QMessageBox.warning(self, "Ошибка экспорта в Excel",
                                        f"Произошла ошибка при создании Excel файла:\n{str(e)}")

        return convert_callback

    def on_download_progress(self, received, total):
        if total > 0:
            progress = int(received * 100 / total)
            self.status_bar.showMessage(f"Скачивание: {progress}%")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())