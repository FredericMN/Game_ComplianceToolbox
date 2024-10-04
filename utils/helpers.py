# 示例辅助函数

def center_window(window):
    """将窗口居中显示"""
    desktop = window.app.primaryScreen().availableGeometry()
    w, h = desktop.width(), desktop.height()
    window.move((w - window.width()) // 2, (h - window.height()) // 2)
