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
	touch "/tmp/$NAME.csv"
	echo "$(pwd)"
	echo "$(pwd)/$NAME.$EXT"
	FILTER='calc8'
	[ 'fods' = "${EXT,,}" ] && FILTER='OpenDocument Spreadsheet Flat XML'
	[ 'tsv' = "${EXT,,}" ] && FILTER='Text - txt - csv (StarCalc)'
	soffice --convert-to "$EXT:$FILTER" --outdir "$(pwd)" /tmp/$NAME.csv 
#	soffice --convert-to "$EXT:calc8" --outdir "$(pwd)" /tmp/$NAME.csv 
#            --convert-to "html:XHTML Writer File:UTF8" *.doc
#            --convert-to pdf:writer_pdf_Export --outdir /home/user *.doc
#	soffice --convert-to "$EXT:/tmp/$NAME.csv" --outdir "$(pwd)"
#	soffice --convert-to "$EXT" "$NAME.$EXT" --outdir "$(pwd)"
#	soffice --convert-to "$EXT" "$NAME.$EXT"
	rm "/tmp/$NAME.csv"
}
Run "$@"
