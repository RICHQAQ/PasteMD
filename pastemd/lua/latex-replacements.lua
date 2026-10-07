-- 配置表：将常用的正则模式映射为目标替换字符串
-- 优势：方便扩展，只需在 mappings 中添加新项即可
local mappings = {
  -- 匹配 \kern 或 {\kern} 后跟数值和单位 (pt, em, cm, mm, ex, bp)
  -- 模式解释：\\kern%s*%-?%d*%.?%d+%a%a
  {
    pattern = "{\\kern%s*[^}]+}",
    replacement = "\\qquad"
  },
  {
    pattern = "\\kern%s*%-?%d*%.?%d+%a%a",
    replacement = "\\qquad"
  },

  -- LaTeX 2.09 旧式命令改写为 texmath（DOCX/OMML 输出）支持的等价命令。
  -- texmath 解析失败时 Pandoc 会把整个公式按原始 TeX 文本渲染进文档。
  -- \cfrac → \dfrac（结构相同，逐参数分数）
  {
    pattern = "\\cfrac",
    replacement = "\\dfrac"
  },
  -- \mathnormal → \mathit
  {
    pattern = "\\mathnormal",
    replacement = "\\mathit"
  },
  -- {\cal X} → {\mathcal{X}}（保留外层分组括号：{\bf X} 位于其他命令的参数内时，
  -- 吃掉括号会把 \tilde{\bf X} 变成 \tilde\mathbf{X}，texmath 解析失败，
  -- 整条公式按原始 TeX 文本降级渲染；%f 断言保证 \cal 后不是字母，避免误配 \calX 等）
  {
    pattern = "{%s*\\cal%f[^%a]%s*(%a+)%s*}",
    replacement = "{\\mathcal{%1}}"
  },
  -- {\cal {X}} → {\mathcal{X}}（声明 + 花括号分组参数）
  {
    pattern = "{%s*\\cal%f[^%a]%s*(%b{})%s*}",
    replacement = "{\\mathcal{%1}}"
  },
  -- \cal{X} → \mathcal{X}（声明 + 分组参数）
  {
    pattern = "\\cal%f[^%a]%s*({[^{}]-})",
    replacement = "\\mathcal%1"
  },
  -- \cal X / \cal\alpha → \mathcal{X} / \mathcal{\alpha}（声明 + 单 token/命令参数）
  {
    pattern = "\\cal%f[^%a]%s*([%a\\][%a]*)",
    replacement = "\\mathcal{%1}"
  },
  -- {\bf X} → {\mathbf{X}}（括号语义同上）
  {
    pattern = "{%s*\\bf%f[^%a]%s*(%a+)%s*}",
    replacement = "{\\mathbf{%1}}"
  },
  -- {\bf {X}} → {\mathbf{{X}}}（声明 + 花括号分组参数）
  {
    pattern = "{%s*\\bf%f[^%a]%s*(%b{})%s*}",
    replacement = "{\\mathbf{%1}}"
  },
  -- \bf{X} → \mathbf{X}
  {
    pattern = "\\bf%f[^%a]%s*({[^{}]-})",
    replacement = "\\mathbf%1"
  },
  -- \bf \mathcal{X} → \mathbf{\mathcal{X}}（\bf 后跟带参命令：整体捕获，
  -- 须置于 \cal 规则之后，防止先行生成的 \mathcal 被误认为 \bf 的参数）
  {
    pattern = "\\bf%f[^%a]%s*(\\%a+%s*%b{})",
    replacement = "\\mathbf{%1}"
  },
  -- \bf X / \bf\alpha → \mathbf{X} / \mathbf{\alpha}
  {
    pattern = "\\bf%f[^%a]%s*(\\%a+)",
    replacement = "\\mathbf{%1}"
  },
  {
    pattern = "\\bf%f[^%a]%s*([%a]+)",
    replacement = "\\mathbf{%1}"
  },

  -- 示例：在此处添加更多扩展规则
  -- { pattern = "\\mbox%s*(%b{})", replacement = "\\text%1" },
}

--- 核心处理逻辑：遍历配置表进行文本替换
local function apply_replacements(content)
  for _, rule in ipairs(mappings) do
    content = content:gsub(rule.pattern, rule.replacement)
  end
  return content
end

--- Pandoc 过滤器入口
return {
  {
    Math = function(el)
      el.text = apply_replacements(el.text)
      return el
    end
  }
}