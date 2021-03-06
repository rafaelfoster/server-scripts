
#!/usr/bin/python

import os
import re
import grp
import pwd
import gzip
import math
import time
import hashlib
import smtplib
import datetime
from subprocess import Popen, PIPE

# Variable definitions
db_username  = "user_with_access_to_db"
db_password  = "a_very_secure_password"
db_hostname  = "db-server.example.com"
db_bkp_path  = "/var/backups/mysql/%s" % (db_hostname)
db_exeption  = "Database" # To add more Databases to filter (it means, to not make backup of it) use the pipe '|' and the name of another db. Ex: Database|mysql|dbname

# SMTP Settings
smtp_sender      = "bkp-mysql@example.com"
smtp_receivers   = "destination@example.com" # or a list ['user1@example.com', 'user2@example.com']
#smtp_auth_user   = 'username@domain.com.br'
#smtp_auth_passwd = 'us3r_p@ss0rd'
smtp_server_addr = "smtp.example.com"
smtp_server_port = "25"
smtp_server_tls  = False # or True (with the first capital letter)

# User and group that will owner the Backup file
fUserid  = "user"
fGroupid = "group"

def main():

	if not os.path.exists(db_bkp_path):
		os.makedirs(db_bkp_path)

	os.chdir(db_bkp_path)

	cmd_mysql_listdb = ['mysql','-h', db_hostname, '-u', db_username, '-p' + db_password, '-e', 'SHOW DATABASES;']
	p1 = Popen(cmd_mysql_listdb, stdout=PIPE, stderr=PIPE)
	p2 = Popen(['grep','-viE', db_exeption], stdin=p1.stdout, stdout=PIPE, stderr=PIPE)
	dbs = p2.communicate()[0]

	p1_error = p1.stderr.read()

	if not dbs:
		if len(p1_error):
			strError = p1_error
		elif len(p2.stderr.read()):
			p2_error = p2.stderr.read()
			strError = p2_error
		else:
			strError = "Database list is empty."

		sendMail("Execution error: %s" % (strError) )
		exit()

	dblist = re.split("\s+", dbs)
	jobStartTime = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')

	msg = '''<p>
	<b> 
		Job Start time: 
	</b>
	%s 
	</p>

	<table border='1' style='border: 1px solid gray;'>
		<tr align='center' style='background-color: #D9D9D9;'>
			<td>
				<b>Database</b>
			</td>
			<td>
				<b>Size</b>
			</td>
			<td>
				<b>SHA1</b>
			</td>
			<td>
				<b>Time</b>
			</td>
		</tr> 
''' % (db_hostname, jobStartTime)
		
	for dbname in dblist:
		if dbname: 
			fdbname_zip = dbname + ".sql.gz"
			strBkpStart = int(time.time())
			cmd_mysqldump = "mysqldump --single-transaction --routines --quick -h %s -u %s -p%s -B %s " % (db_hostname, db_username, db_password, dbname)

			if os.path.exists(fdbname_zip):
				os.rename(fdbname_zip, fdbname_zip + ".old")

			try:
				oFid = gzip.open(fdbname_zip, 'wb')
				sort = Popen(cmd_mysqldump, shell=True, stdout=PIPE)
				oFid.writelines(sort.stdout)
				oFid.close()
				if os.path.exists(fdbname_zip + ".old"):
					os.remove( fdbname_zip + ".old" )

				if os.path.exists(fdbname_zip):
					fgetsize = os.path.getsize(fdbname_zip)
					uid = pwd.getpwnam(fUserid).pw_uid
					gid = grp.getgrnam(fGroupid).gr_gid
					os.chown(fdbname_zip, uid, gid )
					dbsha1sum = getHash(fdbname_zip)
					
				else:
					msg = msg + "<tr><td> %s </td><td colspan='2'>  Backup file not found! </td></tr>" % (dbname)
					continue

				dbsize = convertSize(fgetsize)
				strBkpEnd = int(time.time())
				d = divmod(strBkpEnd - strBkpStart, 86400)
				h = divmod(d[1],3600)
				m = divmod(h[1],60)
				s = m[1]
				timespend = '%d hours, %d minutes, %d seconds' % (h[0],m[0],s)

				msg = msg + """
					<tr>
						<td> %s </td>
						<td align='right'> %s </td>
						<td> %s </td>
						<td> %s </td>
					</tr>
					""" % (dbname, dbsize, dbsha1sum, timespend)

			except IOError as err:
				print "Error: %s" % (err)
				os.remove( fdbname_zip )
				os.rename(fdbname_zip + ".old", fdbname_zip)
				continue		

	jobEndTime = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
	msg = msg + '''
			</table>
			  <p>
				<b> 
					Job Finished: 
				</b>
				%s 
			</p>
		''' % (jobEndTime)

	sendMail(msg)

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

def sendMail(body):

	header = '''From: Backup Mysql <%s>
To: <%s>
MIME-Version: 1.0
Subject: [Backup Mysql] - Mysql backup report
Content-type: text/html
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body>
  <b>Server Name:</b> %s <br><br>

''' % (smtp_sender, smtp_receivers, db_hostname)

	footer = "</body></html>"

	MailMsg = '''%s
%s
%s 
''' % (header, body, footer)

	print MailMsg

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
		print "Successfully sent email"

	except Exception as e:
		print "Error: unable to send email:  %s" % (e)

# Execute main function
if __name__ == "__main__":
    main()
