"""Application constants."""

# 触发防抖时间（秒）
FIRE_DEBOUNCE_SEC = 0.5

# 重试相关
WORD_INSERT_RETRY_COUNT = 3
WORD_INSERT_RETRY_DELAY = 0.3  # 秒

# 默认通知超时时间
NOTIFICATION_TIMEOUT = 3

# 多文件分开输出时，超过该数量不自动打开（只保存），避免一次弹出大量窗口
BATCH_OPEN_LIMIT = 5

# 清理等待时间
CLEANUP_DELAY = 1.0  # 秒

# 缓存删除相关
DEFAULT_DELETE_RETRY = 3
DEFAULT_DELETE_WAIT  = 0.05

# 剪贴板 HTML (CF_HTML) 轮询读取相关（毫秒）
CLIPBOARD_HTML_WAIT_MS = 500
CLIPBOARD_POLL_INTERVAL_MS = 20
