import socket, ssl, email,traceback,os,platform,logging,time

# Only tested with Python 3.6.1 on Windows.  It didn't work with 3.5 running on Linux.

# Login Info:
USERNAME='ImapUsername'
PASSWORD='Password'
IMAPSERVER='imapserver.com'
IMAPPORT=993

# Having problems?  Set the following to logging.DEBUG
LOGLEVEL=logging.INFO

# RFC says the server can disconnect at 29 minutes so set this
# to something less than that.
IDLE_TIME=15*60

# The string prefixing every command that is sent.  Just leave it if
# you have no reason to change it.
cmdPrefix='abc ' # Must end with a space.

# Change svdir to where you want your downloaded files and log to be saved.
if platform.system()=='Windows':
	isWindows=True
	svdir='c:\\faxlog\\'
	import win32api
	import win32print
else: # Linux
	isWindows=False
	svdir='/tmp/'
	import subprocess

# Script logs at logging.INFO or logging.DEBUG.  DEBUG produces a lot of
# output so don't run it at that for long periods of time.
logging.basicConfig(level=LOGLEVEL,filename=os.path.join(svdir,'faxprinter.log'),format='%(asctime)s-%(name)s-%(levelname)s: %(message)s')
log = logging.getLogger('faxprint')

#
# No configuration changes required beneath this point.
#

s  = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
ss = ssl.wrap_socket(s)
fread=ss.makefile()

def sendcmd(cmd):
	global ss
	if cmd!='DONE':
		cmd=cmdPrefix+cmd
	cmd=cmd+'\r\n'
	ss.sendall(cmd.encode('utf-8'))


def connect():
	global s,ss
	log.info('Connecting...')
	s  = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
	ss = ssl.wrap_socket(s)
	ss.setblocking(True)
	ss.settimeout(IDLE_TIME)
	while True:
		try:
			ss.connect((IMAPSERVER,IMAPPORT))
			break
		except:
			disconnect()
	tmp=receiveall()
	log.debug(tmp)
	sendcmd('LOGIN %s %s' % (USERNAME,PASSWORD))
	tmp=receiveall()
	log.debug(tmp)
	log.info('Logged in.')
	selectmailbox('INBOX')


def disconnect():
	global ss,fread
	fread.close()
	time.sleep(5)
	ss.shutdown(socket.SHUT_RDWR)
	time.sleep(5)
	ss.close()
	time.sleep(5)
	log.info("Disconnected.")
	

def selectmailbox(mailboxname):
	global ss
	sendcmd('SELECT %s' % (mailboxname,))
	tmp=receiveall()
	log.debug(tmp)
	if ' OK ' not in tmp:
		log.info("Didn't get a good response from selecting the mailbox.")



# Receive responses in up to 4k blocks, the recommended size in the Python docs.
def receiveall():
	global ss
	whole=b''
	while True:
		chunk=ss.recv(4096)
		whole=whole+chunk
		if len(chunk)<4096:
			log.debug('Chunk is %d, so exiting.' % (len(chunk),))
			break
	return whole.decode('utf-8')


# Returns a list of all the unread message ids.
# Currently only used in getnewmessagesandprint()
def getnewmessageids():
	sendcmd('SEARCH UNSEEN')
	tmp=receiveall()
	tmp=tmp.split('\r\n')
	msgids=[]
	if len(tmp)>=2 and 'OK' in tmp[1]:
		tmp=tmp[0].split(' ')
		for msgid in tmp:
			if msgid in ('*','SEARCH'): continue
			msgids.append(msgid)

	return msgids


# This function iterates through all the unread msgids and if it
# finds an attachment, it sends it to the printer.
# I highly recommend you validate the 'From' field so you
# don't end up printing a bunch of spam.
def getnewmessagesandprint():
	msgids=getnewmessageids()
	for msgid in msgids:
		cmd='FETCH %s RFC822' % (msgid,)
		sendcmd(cmd)
		# The commented bit doesn't work.  For some reason,
		# grabbing the entire response as one big blob and
		# then splitting it into lines to process it doesn't
		# work.  Instead, I have read the response line by
		# line.  If anyone has any bright ideas as to why that
		# is, please let me know.
		'''
		body=receiveall()
		body=body.split('\r\n')
		linecount=0
		for line in body:
			linecount=linecount+1
			if 'FETCH (RFC822' in line: body.remove(line)
			elif line==')': body.remove(line)
			elif line[:6]==cmdPrefix+'OK': body.remove(line)
			elif line.strip()[:4]=='FLAG': body.remove(line) # Outlook sticks this in.
		body='\r\n'.join(body)
		log.debug(body)
		log.debug("There are %d lines" % (linecount,))
		'''
		body=''
		while True:
			line=fread.readline()
			if 'FETCH (RFC822' in line: continue
			elif line==')\r\n': continue
			elif line.strip()[:4]=='FLAG': continue # Outlook.com sticks this in.
			elif line[:6]==cmdPrefix+'OK': break
			log.debug(line)
			body=body+line
		e=email.message_from_string(body) # .decode('utf8'))
		# Make sure the email comes from our fax provider.
		if e.__contains__('From') and e['From'].find('onlinefaxes.com')>-1:
			for part in e.walk():
				if part.get_content_maintype() == 'multipart':
					continue
				if part.get('Content-Disposition') is None:
					continue

				filename=part.get_filename()
				if filename is not None:
					sv_path = os.path.join(svdir, filename)
					fp = open(sv_path, 'wb')
					fp.write(part.get_payload(decode=True))
					fp.close()
					if isWindows:
						# Now print it.
						win32api.ShellExecute ( 0, "print", sv_path, '/d:"%s"' % win32print.GetDefaultPrinter (), ".", 0)
					else:
						subprocess.run('/usr/bin/lpr',input=part.get_payload(decode=True))


# This is whole point of this script.  It blocks until a new email arrives or timeouts and
# then issues a DONE allowing getnewmessagesandprint() to look for new emails.
def Idle():
	sendcmd('IDLE')
	tmp=receiveall()
	log.debug('After IDLE Call:'+tmp+':end After IDLE Call')
	if tmp.upper()[:5] != '+ IDL':     # Postfix says '+ idling' but outlook.com says '+ IDLE accepted...'
		log.info("Couldn't start IDLE")  # so this should work for those two anyway...
	else:
		log.info("Idling")
	tmp=receiveall()
	log.debug('After wakeup from idle:'+tmp+':end wakeup')
	
	sendcmd('DONE')
	tmp=receiveall()
			
	log.debug('After DONE call:'+tmp+':end DONE call')


connect()

while True:
	try:
		getnewmessagesandprint()
		Idle()
	except socket.timeout: # IDLE timed out without receiving an email.  Just send DONE and start again.
		try:
			sendcmd('DONE')
			tmp=receiveall()
			log.debug(tmp)
		except:
			connect() # This nested exception case is *very* important.
	except KeyboardInterrupt:
		log.info('Caught interrupt signal.')
		sys.exit()
	except ConnectionResetError:
		log.info(traceback.format_exc())
		# disconnect() # <-- Shouldn't need this as the connection is already closed.
		connect()
	except:
		log.info(traceback.format_exc())
		disconnect()
		connect()


