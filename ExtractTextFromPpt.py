import os
import re
import codecs

def main():
	current_path = os.getcwd()
	current_path += "/Stuff/ppt/slides/"
	fileList = os.listdir(current_path)
	outputFile = open(current_path+"logfile.txt",'w')
	fileList = sortFiles(fileList)
	for element in fileList:
		if element.find(".xml")!=-1:
			outputFile.write(parseFile(element,current_path))

def parseFile(xmlFile,current_path):
	fileContents= ""
	parser = re.compile('>([A-z0-9 ,\(\)\-\:\.\n\?\!;]+)<')
	regexResults = ""
	try:
		parseXML = open(current_path+xmlFile,'rb')
		fileContents = parseXML.read().decode("cp1252",'ignore')
		results =  re.findall(parser,fileContents)
		for element in results:
			if(element.find("style")==-1):
				if(len(element)==1):
					regexResults += element
				else:
					regexResults = regexResults + element + "\n"

	except Exception as err:
		print("Oh no! We had an error opening "+xmlFile+" :(")
		print(err)
		print(err.with_traceback)

	return regexResults	

def sortFiles(fileList):
	length = len(fileList)
	i=0
	while i<length:
		if fileList[i].find(".xml")==-1:
			fileList.pop(i)
			length-=1
		else:
			fileList[i]=int(fileList[i].replace("slide","").replace('.xml',""))
			i+=1
	fileList.sort()
	i=0	
	while i<length:
		fileList[i] = "slide"+str(fileList[i])+".xml"
		i+=1

	return fileList

main()