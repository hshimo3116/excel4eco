# Makefile for excel4eco
MAKEFILE_DIR := $(dir $(abspath $(lastword $(MAKEFILE_LIST))))
EXCEL_FILE ?= $(MAKEFILE_DIR)workbook.xlsm
MACRO_DIR ?= lib
POWERSHELL ?= $(MAKEFILE_DIR)bin/powershell.sh

source:
	$(POWERSHELL) bin/extract_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"

install:
	$(POWERSHELL) bin/install_macros.ps1 "$(EXCEL_FILE)" "$(MACRO_DIR)"

edit:
	$(EDITOR) $(MACRO_DIR)/*
