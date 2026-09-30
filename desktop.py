"""Payroll Desk: a self-contained local web application in a native Qt window."""
import json
import os
from pathlib import Path
import secrets
import socket
import sys
import threading
import time
import traceback
import urllib.request

import uvicorn
from PySide6.QtCore import QFile, QIODevice, QObject, QTimer, QUrl, Slot
from PySide6.QtGui import QDesktopServices, QIcon
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineScript, QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox

from payroll_app.api import create_app
from payroll_app.service import Store

ROOT = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))


class Bridge(QObject):
    def __init__(self, window, store):
        super().__init__(window)
        self.window, self.store = window, store

    @Slot(str, str, result=str)
    def choose(self, purpose, suggested):
        try:
            if purpose in ('import', 'restore', 'logo'):
                filters = {'import': 'Excel workbook (*.xlsx)', 'restore': 'Payroll database (*.db)', 'logo': 'PNG image (*.png)'}
                path, _ = QFileDialog.getOpenFileName(self.window, 'Choose ' + purpose, '', filters[purpose])
            elif purpose in ('export', 'backup'):
                name = Path(suggested).name
                suffix = Path(name).suffix.lower()
                filters = {'.pdf': 'PDF document (*.pdf)', '.xlsx': 'Excel workbook (*.xlsx)', '.docx': 'Word document (*.docx)', '.db': 'Payroll database (*.db)'}
                if suffix not in filters:
                    raise ValueError('Unsupported output format.')
                path, _ = QFileDialog.getSaveFileName(self.window, 'Save ' + purpose, str(Path.home() / 'Documents' / name), filters[suffix])
                if path and not Path(path).suffix:
                    path += suffix
            else:
                raise ValueError('Unsupported file action.')
            return json.dumps({'ticket': self.store.ticket(path, purpose), 'name': Path(path).name} if path else {'cancelled': True})
        except Exception as error:
            return json.dumps({'error': str(error)})

    @Slot(str, result=bool)
    def openSaved(self, ticket):
        try:
            path = self.store.consume(ticket, 'open')
            return QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        except ValueError:
            return False


class LocalPage(QWebEnginePage):
    def __init__(self, origin, parent):
        super().__init__(parent)
        self.origin = origin

    def acceptNavigationRequest(self, url, nav_type, main_frame):
        return url.toString().startswith(self.origin + '/') or url.toString() == 'about:blank'


class Window(QMainWindow):
    def __init__(self, origin, secret, store):
        super().__init__()
        self.setWindowTitle('Payroll Desk')
        self.setMinimumSize(800, 480)
        screen = QApplication.primaryScreen().availableGeometry()
        self.resize(min(1440, max(800, screen.width()-60)), min(900, max(480, screen.height()-60)))
        self.view = QWebEngineView(self)
        page = LocalPage(origin, self.view)
        self.view.setPage(page)
        self.setCentralWidget(self.view)
        page.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, False)
        page.settings().setAttribute(QWebEngineSettings.WebAttribute.JavascriptCanOpenWindows, False)
        self.channel = QWebChannel(page)
        self.bridge = Bridge(self, store)
        self.channel.registerObject('desktop', self.bridge)
        page.setWebChannel(self.channel)
        source = QFile(':/qtwebchannel/qwebchannel.js')
        if not source.open(QIODevice.OpenModeFlag.ReadOnly):
            raise RuntimeError('Qt WebChannel resources are missing. Reinstall Payroll Desk.')
        script = QWebEngineScript()
        script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentCreation)
        script.setWorldId(QWebEngineScript.ScriptWorldId.MainWorld)
        script.setRunsOnSubFrames(False)
        script.setSourceCode(bytes(source.readAll()).decode('utf-8') + '\n' +
            'window.payrollToken=' + json.dumps(secret) + ';' +
            'window.nativeReady=new Promise(resolve=>new QWebChannel(qt.webChannelTransport,c=>{window.desktop=c.objects.desktop;resolve(window.desktop)}));')
        page.scripts().insert(script)
        self.view.load(QUrl(origin + '/'))

    def foreground(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def closeEvent(self, event):
        if getattr(self, '_close_confirmed', False):
            event.accept()
            return
        event.ignore()
        def checked(state):
            dirty, busy = state if isinstance(state, list) else (False, False)
            if busy:
                QMessageBox.information(self, 'Operation in progress', 'Please wait for the current payroll operation to finish before closing.')
                return
            if dirty and QMessageBox.question(self, 'Unsaved changes', 'Discard unsaved changes and close Payroll Desk?') != QMessageBox.StandardButton.Yes:
                return
            self._close_confirmed = True
            self.close()
        self.view.page().runJavaScript('[Boolean(window.payrollDirty),Boolean(window.payrollBusy)]', checked)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName('PayrollDesk')
    app.setOrganizationName('PayrollDesk')
    data_dir = os.environ.get('PAYROLL_DATA_DIR')
    name = 'PayrollDesk-' + __import__('hashlib').sha256(str(data_dir or os.environ.get('LOCALAPPDATA', '')).encode()).hexdigest()[:16]
    client = QLocalSocket()
    client.connectToServer(name)
    if client.waitForConnected(500):
        client.write(b'activate')
        client.waitForBytesWritten(500)
        return 0
    QLocalServer.removeServer(name)
    local_server = QLocalServer()
    if not local_server.listen(name):
        QMessageBox.critical(None, 'Payroll Desk', 'Another copy is starting. Please try again.')
        return 1
    backend = None
    thread = None
    try:
        legacy = Path(os.environ.get('PAYROLL_LEGACY_DB') or (Path(sys.executable).parent / 'payroll.db' if getattr(sys, 'frozen', False) else ROOT / 'payroll.db'))
        store = Store(data_dir, legacy)
        assets = ROOT / 'frontend' / 'out'
        if not (assets / 'index.html').exists():
            raise RuntimeError('The interface build is missing. Run scripts\\build.ps1 or reinstall the application.')
        secret = secrets.token_urlsafe(32)
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.bind(('127.0.0.1', 0))
        port = sock.getsockname()[1]
        origin = f'http://127.0.0.1:{port}'
        backend = uvicorn.Server(uvicorn.Config(create_app(store, secret, assets), host='127.0.0.1', port=port, log_config=None, access_log=False))
        thread = threading.Thread(target=backend.run, kwargs={'sockets': [sock]}, daemon=True)
        thread.start()
        for _ in range(200):
            try:
                req = urllib.request.Request(origin + '/api/health', headers={'X-Payroll-Token': secret})
                with urllib.request.urlopen(req, timeout=0.2) as response:
                    if response.status == 200:
                        break
            except Exception:
                if not thread.is_alive():
                    raise RuntimeError('Local payroll service stopped during startup.')
                time.sleep(0.05)
        else:
            raise RuntimeError('Local payroll service did not start. Restart the application and check antivirus restrictions on localhost.')
        window = Window(origin, secret, store)
        def activate():
            connection = local_server.nextPendingConnection()
            if connection:
                connection.disconnectFromServer()
            window.foreground()
        local_server.newConnection.connect(activate)
        window.show()
        # Optional automated rendering check; never used in normal operation.
        if '--smoke' in sys.argv:
            target = Path(sys.argv[sys.argv.index('--smoke') + 1])
            def capture():
                window.view.page().runJavaScript('window.scrollTo(0,0)')
                def save_capture():
                    window.grab().save(str(target))
                    window.view.page().runJavaScript('JSON.stringify({title:document.title,text:document.body.innerText,ready:!!window.desktop})', lambda result: (target.with_suffix('.json').write_text(str(result), encoding='utf-8'), app.quit()))
                QTimer.singleShot(300, save_capture)
            QTimer.singleShot(12000, capture)
        return app.exec()
    except Exception:
        message = traceback.format_exc()
        directory = Path(data_dir or Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'PayrollDesk')
        directory.mkdir(parents=True, exist_ok=True)
        (directory / 'startup-error.log').write_text(message, encoding='utf-8')
        QMessageBox.critical(None, 'Payroll Desk could not start', message.splitlines()[-1] + '\n\nDetails: ' + str(directory / 'startup-error.log'))
        return 1
    finally:
        if backend:
            backend.should_exit = True
        if thread:
            thread.join(timeout=5)
        local_server.close()


if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    sys.exit(main())
