"""Tray icon generation and management."""

import os
from PIL import Image, ImageDraw

from ...config.paths import get_app_png_path


def create_fallback_icon(ok: bool = True, flash: bool = False) -> Image.Image:
    """
    创建备用图标（当无法加载资源图标时使用）
    
    Args:
        ok: 是否为正常状态
        flash: 是否为闪烁效果
        
    Returns:
        PIL 图像对象
    """
    size = (64, 64)
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # 背景
    bg_color = (30, 30, 30, 255)
    draw.rectangle([0, 0, size[0], size[1]], fill=bg_color)
    
    # 状态色彩
    color = (60, 200, 80, 255) if ok else (220, 70, 70, 255)
    if flash:
        color = tuple(min(255, int(c * 1.3)) if i < 3 else c for i, c in enumerate(color))
    
    # 绘制圆形
    draw.ellipse([10, 10, 54, 54], fill=color)
    
    return img


def load_base_icon() -> Image.Image:
    """
    加载基础图标
    
    Returns:
        PIL 图像对象
    """
    try:
        icon_path = get_app_png_path()
        if os.path.exists(icon_path):
            return Image.open(icon_path).convert("RGBA")
    except Exception:
        pass
    
    # 回退到生成的图标
    return create_fallback_icon(ok=True)


def create_status_icon(ok: bool) -> Image.Image:
    """
    创建带状态指示的托盘图标
    
    Args:
        ok: 是否为正常状态
        
    Returns:
        PIL 图像对象
    """
    base = load_base_icon().copy()
    width, height = base.size
    draw = ImageDraw.Draw(base)
    
    # 计算状态角标的位置和大小
    radius = int(min(width, height) * 0.28)
    padding = int(radius * 0.25)
    
    # 右下角位置
    x1 = width - radius - padding
    y1 = height - radius - padding
    x2 = width - padding
    y2 = height - padding
    
    # 绘制白色边框
    draw.ellipse([x1 - 2, y1 - 2, x2 + 2, y2 + 2], fill=(255, 255, 255, 255))
    
    # 绘制状态色彩
    status_color = (60, 200, 80, 255) if ok else (220, 70, 70, 255)
    draw.ellipse([x1, y1, x2, y2], fill=status_color)
    
    return base
