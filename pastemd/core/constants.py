"""Application constants."""

# 触发防抖时间（秒）
FIRE_DEBOUNCE_SEC = 0.5

# 重试相关
WORD_INSERT_RETRY_COUNT = 3
WORD_INSERT_RETRY_DELAY = 0.3  # 秒

# 默认通知超时时间
NOTIFICATION_TIMEOUT = 3

# 清理等待时间
CLEANUP_DELAY = 1.0  # 秒

# 缓存删除相关
DEFAULT_DELETE_RETRY = 3
DEFAULT_DELETE_WAIT  = 0.05
