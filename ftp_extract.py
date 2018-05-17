import pandas, urllib, os, zipfile, sys, datetime
import easygui
import atexit

def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

def queryData(df):

	before = len(df)

	queryfield = easygui.choicebox("Which field would you like to query?", "Query Data", df.columns.values.tolist())
	
	relations = ["x > y", "x < y", "x >= y", "x <= y", "x = y", "x != y"]
	relation = easygui.choicebox("What relation are you looking for?", "Query Data", relations)
	
	value = easygui.enterbox("What value are you looking for in this query? (Ex. {} where y is ???)".format(relation))
	if(value.isdigit()):
		value = int(value)
	elif(isfloat(value)):
		value = float(value)
	
	if(relation == relations[0]):
		data = df.loc[df[queryfield] > value]
	elif(relation == relations[1]):
		data = df.loc[df[queryfield] < value]
	elif(relation == relations[2]):
		data = df.loc[df[queryfield] >= value]
	elif(relation == relations[3]):
		data = df.loc[df[queryfield] <= value]
	elif(relation == relations[4]):
		data = df.loc[df[queryfield] == value]
	elif(relation == relations[5]):
		data = df.loc[df[queryfield] != value]
	
	after = len(data)
	
	easygui.msgbox("{} records have been selected from the original set of {} records.".format(after, before), "Query Results")
	
	return data
		

def doCsv(filepath):
	df = pandas.read_csv(filepath)
	
	URLfield = easygui.choicebox("Which field contains the FTP path?", "CSV Download", df.columns.values.tolist())
	
	if(easygui.ynbox('Would you like to execute a simple query on this data? (If "No", all {} values will be extracted.)'.format(URLfield))):
		df = queryData(df)
	
	doDownload(df[URLfield].tolist())

def doExcel(filepath):
	df = pandas.read_excel(filepath)
	
	URLfield = easygui.choicebox("Which field contains the FTP path?", "Excel File Download", df.columns.values.tolist())
	
	if(easygui.ynbox('Would you like to execute a simple query on this data? (If "No", all {} values will be extracted.)'.format(URLfield))):
		df = queryData(df)
	
	doDownload(df[URLfield].tolist())

def doText(filepath):
	lines = open(filepath).readlines()
	df = []
	for line in lines:
		df.append(line[:-1])
	doDownload(df)

def doErrorLog(errorlog, incomplete):
	if(errorlog):
		now = datetime.datetime.now()
		outfile = easygui.filesavebox("Where would you like the log?", "Incomplete Downloads Log", "Incomplete_FTP_" + now.strftime("%Y-%m-%d@%H%M") + ".txt")
		with open(outfile, "w+") as file:
			file.write("\n".join(incomplete))
		atexit.unregister(doErrorLog)
		atexit.register(doErrorLog, False, [])
	
def doDownload(pathlist):
	print(pathlist)
	print(len(pathlist))
	unzip = False
	incomplete = pathlist
	for item in pathlist:
		if(item[item.rfind("."):] == ".zip"):
			unzip = easygui.ynbox("We have detected at least one zipped file in this list. Would you like all zipped items to be unzipped?", "Unzip Files?")
			break
	outdir = easygui.diropenbox(title="Save FTP Files to Directory")
	easygui.msgbox("Downloading of {} files will start when this message is dismissed. Progress can be monitored on the command prompt window. You must keep the command prompt window open to keep the program running.".format(len(pathlist)))
	
	i = 0
	for line in pathlist:
	
		print(i)
		i = i + 1
		print(line)
		
		saveas = outdir + "\\" + line[line.rfind("/")+1:].replace("\\","") if "/" in line else outdir + "\\" + line.replace("\\","")
		print("\n" + datetime.datetime.now().strftime("%Y-%m-%d at %H:%M") + "\nDownloading {}".format(line))
		try:
			urllib.urlretrieve(line, saveas)
			incomplete.remove(line)
			if(len(incomplete) > 0):
				atexit.unregister(doErrorLog)
				atexit.register(doErrorLog, True, incomplete)
			
			print("Saved file as " + saveas)
			
			if(unzip and item[item.rfind("."):] == ".zip"):
				print("Extracting files from {}".format(saveas))
				zip_ref = zipfile.ZipFile(saveas, 'r')
				zip_ref.extractall(outdir)
				zip_ref.close()
				os.remove(saveas)
				print("Extracted {} to {}".format(saveas, outdir))
		except IOError:
			print("----------------------------------------------------------------------------------------\n\tCould not access {} at this time!\n----------------------------------------------------------------------------------------".format(line))
		except:
			print "Unexpected error:", sys.exc_info()[0]
			raise
		
	if(len(incomplete) == 0):
		print("Downloaded All Files!")
	else:
		recurse = easygui.ynbox("{} files were not downloaded. Would you like to try to download them again?".format(len(incomplete)), "Incomplete Download")
		if(recurse):
			doDownload(incomplete)
		else:
			errorlog = easygui.ynbox("Would you like to log the files that could not be downloaded? This log is formatted to be used with this program.", "Incomplete Downloads")
			if(errorlog):	
				atexit.unregister(doErrorLog)
				atexit.register(doErrorLog, errorlog, incomplete)
			else:
				atexit.unregister(doErrorLog)
				atexit.register(doErrorLog, errorlog, [])

easygui.msgbox(msg='This tool will automatically download FTP files when provided with a list of FTP paths.', title="FTP Extractor")
infile = easygui.fileopenbox("Select file with FTP paths")
ext = infile[infile.rfind("."):]

if(ext == ".csv"):
	doCsv(infile)
elif(ext == ".xlsx" or ext == ".xls"):
	doExcel(infile)
elif(ext == ".txt"):
	doText(infile)
else:
	easygui.msgbox("The file {} is not of a supported type. Supported types are .csv, .xlsx, .xls, and .txt.".format(infile), "Unsupported File Type")
	sys.exit()
	