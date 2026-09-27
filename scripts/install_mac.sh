#!/bin/sh
set -eu

project_root=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
helper="$project_root/dist/helper/ppt2fig"
support_dir="$HOME/Library/Application Support/PPT2Fig"
script_dir="$HOME/Library/Application Scripts/com.microsoft.Powerpoint"
mkdir -p "$support_dir" "$script_dir"
if [ -x "$helper" ]; then
  cp "$helper" "$support_dir/ppt2fig"
else
  python3 -m venv "$support_dir/runtime"
  "$support_dir/runtime/bin/python" -m pip install "$project_root[gui]"
  cat > "$support_dir/ppt2fig" <<EOF
#!/bin/sh
exec "$support_dir/runtime/bin/ppt2fig" "\$@"
EOF
fi
chmod 755 "$support_dir/ppt2fig"
if ! "$support_dir/ppt2fig" doctor; then
  echo "Poppler 不完整。请安装 Poppler，或设置 PPT2FIG_POPPLER_DIR 指向其 bin 目录。" >&2
  exit 1
fi
osacompile -o "$script_dir/PPT2FigBridge.scpt" "$project_root/office/PPT2FigBridge.applescript"
echo "辅助程序已安装。然后在 PowerPoint 的 工具 → PowerPoint 加载项 中添加 addin/PPT2Fig-v1.ppam。"
