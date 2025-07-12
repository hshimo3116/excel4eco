# Makefile for excel4eco
EXCEL_FILE ?= workbook.xlsm
MACRO_DIR ?= lib

source:
	cscript //nologo bin/extract_macros.vbs $(EXCEL_FILE) $(MACRO_DIR)

install:
	cscript //nologo bin/install_macros.vbs $(EXCEL_FILE) $(MACRO_DIR)

edit:
	$(EDITOR) $(MACRO_DIR)/*
