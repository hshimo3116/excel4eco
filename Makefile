# Makefile for excel4eco
MAKEFILE_DIR := $(dir $(abspath $(lastword $(MAKEFILE_LIST))))
EXCEL_FILE ?= $(MAKEFILE_DIR)workbook.xlsm
MACRO_DIR ?= lib

source:
	cscript //nologo bin/extract_macros.vbs "$(EXCEL_FILE)" "$(MACRO_DIR)"

ps_source:
	powershell -ExecutionPolicy Bypass -File bin/extract_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"

install:
	cscript //nologo bin/install_macros.vbs "$(EXCEL_FILE)" "$(MACRO_DIR)"

ps_install:
	powershell -ExecutionPolicy Bypass -File bin/install_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"

edit:
	$(EDITOR) $(MACRO_DIR)/*
