import os,sys,shutil,time
import binascii
import base64
import win32com.client

global path
path = os.getcwd()
global avPath
global scanCmd
global blkfld
global blkfld1
global counter
counter = 0
global dircnt
dircnt = 1
global exten
global block

et1 = raw_input("Enter File Extension ( eg. .doc .ppt .xls etc): ")
exten = str(et1)
bk1 = raw_input("Enter Block size : ")
block = int(bk1)

################################
''' Set AV path and command '''
################################
def setAVpath():
	os.chdir("c:\\")
	global avPath
	global scanCmd
	for root, dirs, files in os.walk("Program Files"):
		rootPath = "C:\\"+root
		for f in files:
			# print f
			if f == "ashCmd.exe":
				avPath = rootPath
				scanCmd = "ashCmd.exe /e=100 /p=1 \""
			elif f == "avp.com":
				avPath = rootPath
				scanCmd = "avp.com SCAN /i4 /fa \""
			elif f == "bdc.exe":
				avPath = rootPath
				scanCmd = "bdc.exe -del \""
			elif f == "avgscanx.exe":
				avPath = rootPath
				scanCmd = "avgscanx.exe /clean /heur /scan=\""
			elif f == "clamscan.exe":
				avPath = rootPath
				scanCmd = "clamscan.exe --infected --remove=yes \""
			elif f == "AdAwareCommandLineScanner.exe":
				avPath = rootPath
				scanCmd = "AdAwareCommandLineScanner.exe --delete \""
			elif f == "a2cmd.exe":
				avPath = rootPath
				scanCmd = "a2cmd.exe /a /r /h \""
			elif f == "Sav32cli.exe":
				avPath = rootPath
				scanCmd = "sav32cli.exe -REMOVE -di -nc -p=\""
			elif f == "scancl.exe":
				avPath = rootPath
				scanCmd = "scancl.exe --defaultaction=delete \""
			
			elif f == "fsav.exe":
				avPath = rootPath
				scanCmd = "fsav.exe /NOBOOT /DELETE \""
			# elif f == "MpCmdRun.exe":
				# avPath = rootPath
				# scanCmd = "MpCmdRun.exe -scan -ScanType 3 -file \""
			elif f == "ecls.exe":
				avPath = rootPath
				scanCmd = "ecls.exe /adv-heur /action=clean \""
			elif f == "Pavcl.exe":
				avPath = rootPath
				scanCmd = "Pavcl.exe -aex -nob -auto -cmp -heu:3 \""
			# else:
				# print "Antivirus could not be found"
				# raw_input("Press enter to exit...")
				# sys.exit(0)
			# print "avPath : "+avPath
	# raw_input("...")
	# # print "avPath : "+avPath
	# print "scanCmd : "+scanCmd
	# raw_input("...")
	os.chdir(path)


################################
'''  Zero out given Blocks  '''
################################	
def blockSplitter():
	os.chdir(path)
	global blkfld
	for files in os.listdir("."):
		fileName, fileExtension = os.path.splitext(files)
		if fileExtension == exten:
			ext = fileExtension
			end = os.path.getsize(files)
			read = open(files,'rb').read()
			strt = 0
			if end%block == 0:
				rem = end/block
			else:
				rem = (end/block) + 1
				remainder = end%block
			dir = fileName
			blkfld = dir
			try:
					os.mkdir(dir)
			except:
					shutil.rmtree(dir)
					os.mkdir(dir)
			rem1 = rem
			rem2 = end - strt
			for i in range(strt,rem1):
				print "Blocks remaining - " + str(rem)
				rem-=1
				rem2-=1
				if (end%block) != 0:
					if i == (rem1-1):
						temp = read[:strt] + "\x00" * remainder + read[strt+remainder:]
						fname = dir + "\\" + str(strt) + "-" + str(strt+remainder) + ext
						write =  open(fname,'wb')
						write.write(temp)
						strt += block 
					else:
						temp = read[:strt] + "\x00" * block + read[strt+block:]
						fname = dir + "\\" + str(strt) + "-" + str(strt+block) + ext
						write =  open(fname,'wb')
						write.write(temp)
						strt += block 
				else:
					temp = read[:strt] + "\x00" * block + read[strt+block:]
					fname = dir + "\\" + str(strt) + "-" + str(strt+block) + ext
					write =  open(fname,'wb')
					write.write(temp)
					strt += block


################################
'''    Single byte zero     '''
################################
def fileSplitter1(strt,end,blok):
	os.chdir(path)
	global blkfld1
	global dircnt
	for files in os.listdir("."):
		fileName, fileExtension = os.path.splitext(files)
		if fileExtension == str(exten):
			ext = fileExtension
			read = open(files,'rb').read()
			rem = (end-strt)/blok
			dir = fileName + "-" + str(strt) + "-" + str(end) 
			blkfld1 = dir
			try:
					os.mkdir(dir)
			except:
					shutil.rmtree(dir)
					os.mkdir(dir)
			dircnt += 1
			rem1 = rem
			rem2 = end - strt
			for i in range(0,rem):
				print "processing block - " + str(rem)
				print "Blocks remaining - " + str (rem2)
				rem2-=1
				rem-=1
				if read[strt:strt+1] != "\x00":
					temp = read[:strt] + "\x00" * blok + read[strt+blok:]
				else:
					temp = read[:strt] + "\x01" * blok + read[strt+blok:]
				fname = dir + "\\" + str(strt) + "-" + str(strt+blok) + ext
				write =  open(fname,'wb')
				write.write(temp)
				strt += blok


def officeAutomater(fldr):
	
	os.chdir(fldr)
	li = []
	i = 0
	for file in os.listdir("."):
		fileName1, fileExtension1 = os.path.splitext(file)
		if exten == ".ppt" or exten == "pptx":
			li.append(file)
			PptApplication = win32com.client.Dispatch("PowerPoint.Application")
			try:
				PptApplication.Visible = True
			except:
				time.sleep(10)
				PptApplication = win32com.client.Dispatch("PowerPoint.Application")
				PptApplication.Visible = True
			fopen = open("currentFile.txt", 'wb')
			fopen.write(fileName1)
			fopen.close()
			try:
				ppt = PptApplication.Presentations.Open(fldr+"\\"+file)
				print "Checking " + file
				
			except Exception as e:
				print e
				
			try:
				PptApplication.Quit()
			except Exception as e1:
				print e1
				
			flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
			if flag != "":
				print "Calc executed in file : "+ file
				os.system("pause")
			try:
				os.remove(li[i-1])
			except:
				pass
			i+=1
		##########################################################################	
		if exten == ".doc" or exten == "docx":
			li.append(file)
			wordApplication = win32com.client.Dispatch("Word.Application")
			try:
				wordApplication.Visible = True
			except:
				time.sleep(10)
				wordApplication = win32com.client.Dispatch("Word.Application")
				wordApplication.Visible = True
			fopen = open("currentFile.txt", 'wb')
			fopen.write(fileName1)
			fopen.close()
			try:
				word = wordApplication.Documents.Open(fldr+"\\"+file)
				print "Checking " + file
				
			except Exception as e:
				print e
				
			try:
				wordApplication.Quit()
			except Exception as e1:
				print e1
				
			flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
			if flag != "":
				print "Calc executed in file : "+ file
				os.system("pause")
			try:
				os.remove(li[i-1])
			except:
				pass
			i+=1
		
		##############################################################################
		if exten == ".xls" or exten == "xlsx":
			li.append(file)
			excelApplication = win32com.client.Dispatch("Excel.Application")
			try:
				excelApplication.Visible = True
			except:
				time.sleep(10)
				excelApplication = win32com.client.Dispatch("Excel.Application")
				excelApplication.Visible = True
			fopen = open("currentFile.txt", 'wb')
			fopen.write(fileName1)
			fopen.close()
			try:
				excel = excelApplication.Workbooks.OpenOpen(fldr+"\\"+file)
				print "Checking " + file
				
			except Exception as e:
				print e
				
			try:
				wordApplication.Quit()
			except Exception as e1:
				print e1
				
			flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
			if flag != "":
				print "Calc executed in file : "+ file
				os.system("pause")
			try:
				os.remove(li[i-1])
			except:
				pass
			i+=1
			
def storeSignature(fldpath):
	li = []
	for file in os.listdir(fldpath):
		temp,x = file.split("-")
		li.append(int(temp))
	fo = open("SingleByteSig.txt",'ab')
	fo.write(str(li)+"\r\n")
	fo.close()
				
def deleteFolder(fldpath):
	try:
		os.chdir(path)
		shutil.rmtree(fldpath)
	except:
		time.sleep(2)
		shutil.rmtree(fldpath)
				
def singleByteZero():
	global blkfld1
	os.chdir(path+"\\"+blkfld)
	for file in os.listdir("."):
		fileName123, fileExtension123 = os.path.splitext(file)
		f = []
		f = fileName123.split("-")
		b = 1
		s = int(f[0])
		e = int(f[1])
		fileSplitter1(s,e,b)
		scanner(path+"\\"+blkfld1)
		storeSignature(path+"\\"+blkfld1)
		officeAutomater(path+"\\"+blkfld1)
		deleteFolder(path+"\\"+blkfld1)
		

		
def scanner(folder):
	global scanCmd
	os.chdir(avPath)
	os.system(scanCmd + folder + "\"")
	os.chdir(path)
	

	
setAVpath()
blockSplitter()
scanner(path+"\\"+blkfld)
singleByteZero()