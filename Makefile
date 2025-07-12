# Makefile for excel4eco
MAKEFILE_DIR := $(dir $(abspath $(lastword $(MAKEFILE_LIST))))
EXCEL_FILE ?= $(MAKEFILE_DIR)xl/workbook.xlsm
MACRO_DIR ?= src
POWERSHELL ?= $(MAKEFILE_DIR)bin/powershell.sh
EDITOR ?= emacs -nw

# 変数一覧を表示するターゲット
print-vars:
	@echo "Defined Makefile variables:"
	@echo "MAKEFILE_LIST= $(MAKEFILE_DIR)"
	@echo "MAKEFILE_DIR = $(MAKEFILE_DIR)"
	@echo "EXCEL_FILE   = $(EXCEL_FILE)"
	@echo "MACRO_DIR    = $(MACRO_DIR)"
	@echo "POWERSHELL   = $(POWERSHELL)"
	@echo "EDITOR       = $(EDITOR)"

source:
	$(POWERSHELL) bin/extract_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"

install:
	$(POWERSHELL) bin/install_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"

edit:
	$(EDITOR) "$(MACRO_DIR)"/*

register:
	git add "$(MACRO_DIR)"/*.bas "$(MACRO_DIR)"/*.cls "$(MACRO_DIR)"/*.frm
	git commit -m "Update VBA sources"
	# git push

.PHONY: source install edit register
