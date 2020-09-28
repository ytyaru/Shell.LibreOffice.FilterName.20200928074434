#!/usr/bin/env bash
set -Ceu
#---------------------------------------------------------------------------
# LibreOfficeCalcのファイルを新規作成するsofficeコマンド。
# CreatedAt: 2020-09-27
#---------------------------------------------------------------------------
Run() {
	THIS="$(realpath "${BASH_SOURCE:-0}")"; HERE="$(dirname "$THIS")"; PARENT="$(dirname "$HERE")"; THIS_NAME="$(basename "$THIS")"; APP_ROOT="$PARENT";
#	cd "$HERE"
	[ 0 -lt $# ] && { NAME=${1%.*}; EXT=${1#*.}; } || { NAME=$(date +%Y%m%d%H%M%S); EXT=ods; }
#	touch "$csvFile"
	touch "/tmp/$NAME.$EXT"
	echo "$(pwd)"
	soffice --convert-to "$EXT" "$NAME.$EXT" --outdir "$(pwd)"
#	soffice --convert-to "$EXT" "$NAME.$EXT"
	rm "/tmp/$NAME.$EXT"
}
Run "$@"
