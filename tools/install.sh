#!/usr/bin/env bash
# 批量文件处理器 —— 一键安装脚本
# 用法：bash install.sh   或   bash install.sh /path/to/venv
set -euo pipefail

VENV="${1:-$HOME/.venvs/draco-tools}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "==> 创建 Python 环境：$VENV"
[ -d "$VENV" ] || python3 -m venv "$VENV"
PY="$VENV/bin/python"

echo "==> 安装依赖（优先清华源，失败自动换阿里云源）"
if ! "$PY" -m pip install -q --index-url https://pypi.tuna.tsinghua.edu.cn/simple \
      --trusted-host pypi.tuna.tsinghua.edu.cn \
      pymupdf python-docx openpyxl pillow 2>/dev/null; then
  echo "    清华源不可用，改用阿里云源…"
  "$PY" -m pip install -q --index-url https://mirrors.aliyun.com/pypi/simple \
      --trusted-host mirrors.aliyun.com \
      pymupdf python-docx openpyxl pillow
fi

echo "==> 验证安装"
"$PY" -c "
for m in ('fitz','docx','openpyxl','PIL.Image'):
    try:
        __import__(m, fromlist=['*']); print('    OK ' + m)
    except ImportError:
        raise SystemExit('    FAIL ' + m)
"

mkdir -p "$HOME/bin"
printf '#!/usr/bin/env bash\nexec "%s" "%s/batch_tool.py" "$@"\n' "$PY" "$SCRIPT_DIR" > "$HOME/bin/batch-tool"
chmod +x "$HOME/bin/batch-tool"

echo
echo "==> 安装完成 ✅"
echo "    用法： batch-tool --help"
echo "    或：   $PY $SCRIPT_DIR/batch_tool.py --help"
