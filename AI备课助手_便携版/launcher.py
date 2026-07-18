import sys
import os
import time
import threading
import webbrowser
import logging
import traceback

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'launcher_error.log')

try:
    if sys.executable.endswith('pythonw.exe'):
        sys.stdout = open(LOG_FILE, 'w', encoding='utf-8')
        sys.stderr = sys.stdout

    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    if '' not in sys.path:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

    def open_browser():
        time.sleep(2.0)
        webbrowser.open("http://127.0.0.1:5000")

    def main():
        threading.Thread(target=open_browser, daemon=True).start()

        logging.getLogger('werkzeug').setLevel(logging.WARNING)

        import web_app
        web_app.app.run(host='127.0.0.1', port=5000, debug=False)

    if __name__ == '__main__':
        main()
except Exception:
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        traceback.print_exc(file=f)
