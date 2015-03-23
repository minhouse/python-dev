#!/usr/bin/python

import sys, time
from datetime import datetime
from os import system
import subprocess, readline
import socket, netifaces
import re

argvs = sys.argv

def dnscall():
	for line in open('/etc/resolv.conf','r'):
		if re.search(r'^name', line):
			s = line.split()
			return (s[1])

if argvs[1] == "dns":
	print (dnscall())
elif argvs[1] == "status":
	#ip = socket.gethostbyname(socket.gethostname())
	ip = netifaces.ifaddresses('en1')[2][0]['addr']
	print ("IP Address: %s" % ip)
	net = netifaces.ifaddresses('en1')[2][0]['netmask'] 
	print ("Netmask: %s" % net)
	gw = netifaces.gateways()[2][0][0]
	print ("Gateway: %s" % gw)
	dns = dnscall()
	print ("DNS: %s" % dns) 
else:
	print ("Usage: ./os-config.py [ip|host|status|all]")


#hosts = "grep ^Host ~/.ssh/config |awk '{print $2}'"
#hosts = subprocess.getoutput(hosts)
#hosts = hosts.split('\n')

#cmd2 = "/sbin/ifconfig |sed -n 2p"
#cmd2 = subprocess.getoutput(cmd2)

#cmdlist = "/Users/tokuharaminho/shell-lesson/commandlist.txt"

#for h in hosts:
#	for c in open(cmdlist, "r"):
#		ft = c.rstrip()
#		scmd1 = "ssh -n %s %s" % (h,ft)
#		scmd1 = subprocess.getoutput(scmd1)
#		print (scmd1)
