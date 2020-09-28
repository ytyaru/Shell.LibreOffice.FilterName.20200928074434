#!/usr/bin/env bash
set -Ceu
#---------------------------------------------------------------------------
# LibreOfficeのフィルタオプション一覧を取得する。
# CreatedAt: 2020-09-27
#---------------------------------------------------------------------------
Run() {
	THIS="$(realpath "${BASH_SOURCE:-0}")"; HERE="$(dirname "$THIS")"; PARENT="$(dirname "$HERE")"; THIS_NAME="$(basename "$THIS")"; APP_ROOT="$PARENT";
	[ 0 -lt $# ] && { NAME=${1%.*}; EXT=${1#*.}; } || { NAME=$(date +%Y%m%d%H%M%S); EXT=ods; }
	PATH_IN="/tmp/$NAME.csv"
	touch "$PATH_IN"
	GetFilterName() {
		case "${1,,}" in
			'ots') echo 'calc8_template';;
			'fods') echo 'OpenDocument Spreadsheet Flat XML';;
			'csv'|'tsv'|'tab'|'txt') echo 'Text - txt - csv (StarCalc)';;
			'slk'|'sylk') echo 'SYLK';;
			'xlsb') echo 'Calc MS Excel 2007 Binary';;
			'xlsm') echo 'Calc MS Excel 2007 VBA XML';;
			'xlsx') echo 'Calc MS Excel 2007 XML';;
			'xltm'|'xltx') echo 'Calc MS Excel 2007 XML Template';;
			'ods'|*) echo 'calc8';;
#			'pdf') echo 'calc_pdf_Export';; # calc_pdf_addstream_import
#			'jpg'|'jpeg'|'jfif '|'jif '|'jpe') echo 'calc_jpg_Export';;
#			'png') echo 'calc_png_Export';;
		esac
	}
	soffice --convert-to "$EXT:$(GetFilterName "$EXT")" --outdir "$(pwd)" "$PATH_IN"
	rm "$PATH_IN"
}
Run "$@"
