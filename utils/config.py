def format_log_message(message_type, content):
    """统一的日志格式"""
    return f"[{message_type}] {content}"

def format_game_info(info):
    """统一的游戏信息输出格式"""
    return [
        f"  游戏名称：{info[0]}",
        f"  出版单位：{info[1]}",
        # ...其他字段
    ] 