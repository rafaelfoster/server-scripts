#!/usr/bin/python

import os
import re
import grp
import pwd
import sys
import gzip
import math
import time
import yaml
import hashlib
import smtplib
import datetime
from subprocess import Popen, PIPE
import shlex

count = 0
logfile = None
jobname = None
log_file = None
ssh_bin = "/usr/bin/ssh"
mysql_bin = "/usr/bin/mysql"
mysqldump_bin = "/usr/bin/mysqldump"
configFile = "/path/to/config/config.yml"

def main():

	strError = list()
	stream = open(configFile, 'r')
	yml_config = yaml.load(stream)
	

	print yml_config

	for section_key, section_value in yml_config.items():
		global jobname
		jobname = section_key
		cmd_ssh_proxy = None

		tmpdebugVars = list();
		for configkey, configvalue in section_value.items():
			tmpdebugVars.append("Definindo variavel: %s -> %s " % (configkey, configvalue) )
			globals()[configkey] = configvalue
				
		if "log_file_folder" in globals():
			for strToLog in tmpdebugVars:
				_WriteOutput(strToLog)

		# Skip if the section is default
		if jobname == "default": continue;
	
		hostname = db_hostname
		if checkVar('db_ssh_addr'):
			hostname = db_ssh_addr
		dbbkppath = db_bkp_path + "/" + hostname
		
		_WriteOutput("Starting job: " + jobname)
		_WriteOutput("Backup folder: " + db_bkp_path)

		if dbbkppath and not os.path.exists(dbbkppath):
			os.makedirs(dbbkppath)
		os.chdir(dbbkppath)

		if db_hostname and db_username and db_password:
			cmd_mysql_listdb = "%s -h %s -u %s -p%s -e 'SHOW DATABASES;' --skip-column-names" % (mysql_bin, db_hostname, db_username, db_password)

		if checkVar('db_ssh_addr') and checkVar('db_ssh_user'):
			cmd_ssh_proxy = "%s %s@%s" % (ssh_bin, db_ssh_user, db_ssh_addr)
			cmd = '%s "%s"' % ( cmd_ssh_proxy, cmd_mysql_listdb )
			_WriteOutput("Use SSH proxy: " + cmd_ssh_proxy )
		else:
			cmd = cmd_mysql_listdb

		if db_hostname and db_username and db_password:
			p1 = Popen(shlex.split(cmd), stdout=PIPE, stderr=PIPE)
			dbs, p1_error  = p1.communicate()

			if checkVar("db_filter") == True:
				dbsBackup = list()
				for dbfilter in db_filter.split():
					if dbfilter in dbs:
						dbsBackup.append(dbfilter)
				dbs = '\n'.join(dbsBackup)
			elif checkVar("db_exception") == True:
				exception = re.sub(" ","|", db_exception)
				dbs = re.sub( r'%s' % exception, "", dbs)

			if not dbs:
				if p1_error: 
					_WriteOutput("Error: " + p1_error + "\n Database list is empty.")
					strError.append(p1_error)
					strError.append("Database list is empty.")
					break
				sendMail(jobname, "Execution error: %s <br> Error: %s" % (strError, p1_error) )
				continue
		# End If db_hostname

		dblist = re.split("\s+", dbs)

		jobStartTime = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
		msg = '''<p><b> Job Start time: </b>%s </p>
		<table border='1' style='border: 1px solid gray;'>
			<tr align='center' style='background-color: #D9D9D9;'>
				<td><b>Database</b></td>
				<td><b>Size</b></td>
				<td><b>SHA1</b></td>
				<td><b>Time</b></td>
			</tr> 
		''' % (jobStartTime)

		for dbname in dblist:
			if len(dbname) > 0:
				fdbname_zip = dbname + ".sql.gz"

				cmd_mysqldump = "%s --single-transaction --routines --quick -h %s -u %s -p%s -B %s " % (mysqldump_bin, db_hostname, db_username, db_password, dbname)
				if cmd_ssh_proxy:
					cmd = '%s "%s"' % ( cmd_ssh_proxy, cmd_mysqldump )
				else:
					cmd = cmd_mysqldump

				_WriteOutput("Command to Exec: " + cmd)
				strBkpStart = int(time.time())

				if os.path.exists(fdbname_zip):
					_WriteOutput("File already exist, renaming to .old")
					os.rename(fdbname_zip, fdbname_zip + ".old")

				try:
					_WriteOutput("Creating zip file...")
					oFid = gzip.open(fdbname_zip, 'wb')
					sort = Popen(shlex.split(cmd), stdout=PIPE)
					oFid.writelines(sort.stdout)
					oFid.close()
				except IOError as err:
					_WriteOutput("Error: %s" % (err))
					os.remove( fdbname_zip )
					os.rename(fdbname_zip + ".old", fdbname_zip)
					continue
				else:
					_WriteOutput("Writing dump to file: OK")
					if os.path.exists(fdbname_zip + ".old"):
						_WriteOutput("Deleting old backup file...")
						os.remove( fdbname_zip + ".old" )

					if os.path.exists(fdbname_zip):
						fgetsize = os.path.getsize(fdbname_zip)
						uid = pwd.getpwnam(fUserid).pw_uid
						gid = grp.getgrnam(fGroupid).gr_gid
						os.chown(fdbname_zip, uid, gid )
						dbsha1sum = getHash(fdbname_zip)
						_WriteOutput("Backup file size: %s " % (fgetsize) )
					else:
						_WriteOutput("Some error occur. Backup File was not found!")
						msg = msg + "<tr><td> %s </td><td colspan='2'>  Backup file not found! </td></tr>" % (dbname)
						continue

					dbsize = convertSize(fgetsize)
					strBkpEnd = int(time.time())
					d = divmod(strBkpEnd - strBkpStart, 86400)
					h = divmod(d[1],3600)
					m = divmod(h[1],60)
					s = m[1]
					timespend = '%d hours, %d minutes, %d seconds' % (h[0],m[0],s)

					_WriteOutput("Backup Finished after " + timespend)
					msg = msg + '''
						<tr>
							<td> %s </td>
							<td align='right'> %s </td>
							<td> %s </td>
							<td> %s </td>
						</tr>
						''' % (dbname, dbsize, dbsha1sum, timespend)
		# End for dbname

		jobEndTime = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
		msg = msg + '''</table> <p><b> Job Finished: </b> %s </p> ''' % (jobEndTime)
		sendMail(jobname, msg)
		global logfile
		_WriteOutput("\n------------------------\n")
		logfile.close()

		for configkey, configvalue in section_value.items():
			tmpdebugVars.append("Removendo variavel: %s -> %s " % (configkey, configvalue) )
			del globals()[configkey]

	# End For section_key
# End Main()

def convertSize(size):
	size = ( size / 1024 )
	return "%s %s" % (size, "Kb")

def getHash(db_filename):
	BLOCKSIZE = 65536
	hasher = hashlib.sha1()
	with open(db_filename, 'rb') as afile:
		buf = afile.read(BLOCKSIZE)
		while len(buf) > 0:
			hasher.update(buf)
			buf = afile.read(BLOCKSIZE)
	return hasher.hexdigest()

def _WriteOutput(strToWrite):
	global count
	global logfile
	global log_file_folder

	jobDate = datetime.datetime.now().strftime('%Y%m%d')
	log_file = "%s/%s-%s.log" % (log_file_folder, "pyBackupMySql", jobDate)

	if not os.path.exists(log_file_folder):
			os.makedirs(log_file_folder)
	try:
		with open(log_file, "a+") as logfile: 
			if strToWrite:
				dtNow = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
				logLine = "%d) %s - %s" % (count, dtNow, strToWrite)
				if "debug" in map(lambda each:each.lower(), sys.argv):
					print logLine
				if "log_file_folder" in globals():
					try:
						logfile.writelines(logLine)
						logfile.write("\n")
					except IOError as e:
						print "Error while writing log:  %s " % (e)

				count = count + 1

	except IOError as e:
		print "Error: %s " % (e)

def checkVar(varname):
	if varname in globals():
		return True
	else:
		return False

def sendMail(jobname, body ):
	header = '''From: Backup Mysql <bkp-mysql@rodrimar.com.br>
To: <%s>
MIME-Version: 1.0
Subject: [Backup Mysql] - Mysql backup report job: %s
Content-type: text/html
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body>
  <b>Server Name:</b> %s <br><br>

''' % (smtp_receivers, jobname, db_hostname)

	footer = "</body></html>"

	MailMsg = '''%s
%s
%s 
''' % (header, body, footer)

	try:
		smtpObj = smtplib.SMTP(smtp_server_addr, smtp_server_port)

		try:
			smtp_auth_user
		except NameError:
			smtp_auth_required = False
		else:
			if smtp_server_tls:
				smtpObj.ehlo()
				smtpObj.starttls()
				smtpObj.ehlo
			smtpObj.login(smtp_auth_user, smtp_auth_passwd)

		smtpObj.sendmail(smtp_sender, smtp_receivers, MailMsg)
		smtpObj.close()
		_WriteOutput("Successfully sent email")

	except Exception as e:
		_WriteOutput("Error: unable to send email:  %s" % (e))

# Execute main function
if __name__ == "__main__":
    main()
