#
# Simple Makefile for the csv2xlsx project.
#
# Copyright 2022, Stanislas Fechner, stan102@hotmail.fr
#

install:
	@python setup.py install
	@rm -rf dist
	@rm -rf build
	@rm -rf csv2xlsx.egg-info
