#!/usr/bin/env bash
# Xuất PDF bằng Microsoft Word for Mac: tools/render.sh file.docx out.pdf
in="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"; out="$(cd "$(dirname "$2")" && pwd)/$(basename "$2")"
osascript <<OSA
tell application "Microsoft Word"
  set d to open file name (POSIX file "$in" as string)
  save as d file name (POSIX file "$out" as string) file format format PDF
  close d saving no
end tell
OSA
