#!/usr/bin/python

import yaml

stream = open("config.yml", 'r')
yml_config = yaml.load(stream)

for section_key, section_value in yml_config.items():
	print "Sessao: " + section_key
	print ""
	for configkey, configvalue in section_value.items():
		print "%s -> %s" % (configkey, configvalue)
	print ""
	print "----------------------------------------------------"