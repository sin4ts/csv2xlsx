#
# Simple Makefile for the csv2xlsx project.
#
# Copyright 2022, Stanislas Fechner, stan102@hotmail.fr
#

install:
	@ln -fs csv2xlsx.py csv2xlsx
	@python setup.py install
	@rm csv2xlsx
	@rm -rf dist
	@rm -rf build
	@rm -rf csv2xlsx.egg-info
