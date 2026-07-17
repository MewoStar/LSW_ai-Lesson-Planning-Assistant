import sys
sys.path.insert(0, '.')

from web_app import app
import threading
import time

def test_routes():
    time.sleep(1)
    print("\n=== 已注册的路由 ===")
    for rule in app.url_map.iter_rules():
        print(f"  {rule.rule} -> {rule.endpoint}")
    print("=== 路由列表结束 ===\n")

threading.Thread(target=test_routes, daemon=True).start()

if __name__ == "__main__":
    print("""
============================================
       BeiKe Assistant - 可视化生成模块
       http://localhost:6000/module/visualization
============================================
""")
    app.run(host='0.0.0.0', port=6000, debug=False, use_reloader=False, threaded=True)
