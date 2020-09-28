#!/usr/bin/env bash
set -Ceu
#---------------------------------------------------------------------------
# LibreOfficeCalcのフィルタ名を一覧する。
# CreatedAt: 2020-09-27
#---------------------------------------------------------------------------
Run() {
	THIS="$(realpath "${BASH_SOURCE:-0}")"; HERE="$(dirname "$THIS")"; PARENT="$(dirname "$HERE")"; THIS_NAME="$(basename "$THIS")"; APP_ROOT="$PARENT";
	cd "$HERE"
	soffice --calc --norestore "--accept=pipe,name=librepipe;urp;" & {
		# LibreOffice Calc の GUI が起動するまでの時間だけ待機させる必要がある。さもなくば以下エラーになる。
		# __main__.NoConnectException: Connector : couldn't connect to pipe librepipe(10)
		sleep 3
		./filters.py
		jobs
		kill %%
	}
}
Run "$@"
