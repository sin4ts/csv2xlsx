#
# Simple Makefile for the csv2xlsx project.
#
# Copyright 2022, Stanislas Fechner, stan102@hotmail.fr
#

install:
	@pip install .
	@rm -rf dist
	@rm -rf build
	@rm -rf csv2xlsx.egg-info

force-install:
	@pip install . --break-system-packages
	@rm -rf dist
	@rm -rf build
	@rm -rf csv2xlsx.egg-info

uninstall:
	@pip uninstall csv2xlsx

force-uninstall:
	@pip uninstall csv2xlsx --break-system-packages

