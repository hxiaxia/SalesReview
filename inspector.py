import json
import time
import os
from playwright.sync_api import sync_playwright

CONFIG_FILE = "config.json"

def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def run_inspector():
    config = load_config()
    with sync_playwright() as p:
        # 打开 Chrome
        browser = p.chromium.launch(headless=False)
        
        # 挂载我们之前的 Cookie
        state_file = "state.json"
        context_options = {}
        if os.path.exists(state_file):
            context_options["storage_state"] = state_file
            print("已加载您的登录状态！")
            
        context = browser.new_context(**context_options)
        page = context.new_page()

        print(f"打开报表页面: {config['report_url']}")
        page.goto(config['report_url'])
        
        # 注入一段 JS 代码用于捕获您的每次点击并将其外壳打印出来
        print("=" * 60)
        print("💡 页面加载完毕。请您在弹出的浏览器里：")
        print("1. 手动把报表点开（就像您平时操作一样）")
        print("2. 找到那个死活点不到的【项目名称】链接")
        print("3. **不要用右键检查！直接用鼠标左键狠狠点击它一次！**")
        print("=" * 60)
        
        page.add_init_script("""
            document.addEventListener('click', function(e) {
                const element = e.target;
                
                // 向上追溯，把所有的相关信息都捞出来
                let domInfo = "\\n======= 刚刚被点击的元素特征 =======";
                domInfo += "\\n1. 标签名: " + element.tagName;
                domInfo += "\\n2. 类名 (class): " + element.className;
                domInfo += "\\n3. 文本内容: " + element.innerText.substring(0, 50);
                domInfo += "\\n4. 完整 HTML: \\n      " + element.outerHTML.substring(0, 300);
                
                // 找它的父级，通常 a 标签在父级
                if (element.parentElement) {
                     domInfo += "\\n\\n5. 父级标签: " + element.parentElement.tagName;
                     domInfo += "\\n6. 父级完整 HTML: \\n      " + element.parentElement.outerHTML.substring(0, 300);
                }
                domInfo += "\\n======================================\\n";
                
                // 将信息同时打印到浏览器控制台 和 Python 终端
                console.log(domInfo);
                window.playwrightClickedElement(domInfo);
            }, true);
        """)
        
        # 把前端的点击信息传回 Python 控制台
        page.expose_function("playwrightClickedElement", lambda info: print(info))
        
        # 脚本在这个地方无限挂起，等待您去点击页面
        print("\n⏳ 正在监听您的点击...(按 Ctrl+C 退出)")
        page.pause() 
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n监视结束。")
            browser.close()

if __name__ == "__main__":
    run_inspector()
