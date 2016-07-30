import os
import re
import codecs
import zipfile

def main():
	current_path = os.getcwd()
	extractZip(current_path)
	#Loop through all folders that should have powerpoints
	for element in os.listdir(current_path):
		try:
			if((element.find(".")==-1) and (element!="README")):
				processSlides(current_path+"/"+element+"/ppt/slides/")
				#File structure is guaranteed as we just generated it, we can hard code the path in
		except:
			print("Failed on "+current_path+"/"+element)

# Extracts all powerpoints and places the contents in a folder with the same name
# Note: Does write over existing folders
def extractZip(current_path):
	fileList = os.listdir(current_path)
	for element in fileList:
		if(element.find(".pptx")!=-1):
			zippedPpt = zipfile.ZipFile(current_path+"/"+element)
			output_path = current_path +"/"+ element.replace(".pptx","")+"/"
			zippedPpt.extractall(output_path)

def processSlides(current_path):
	fileList = os.listdir(current_path)
	# Splits the file path to write a text file with the powerpoint name to the same directory level as the python file
	outputFile = open(current_path[:current_path.find("/ppt")]+current_path[current_path.rfind("/")+1:]+".txt",'w')
	fileList = sortFiles(fileList)
	for element in fileList:
		if element.find(".xml")!=-1:
			outputFile.write(parseFile(element,current_path))

# Returns all text found in the powerpoint slides
def parseFile(xmlFile,current_path):
	fileContents= ""
	parser = re.compile('>([A-z0-9 \,\(\)\-\:\.\n\?\!\;\'\"]+)<')
	# Trying to handle odd cases, add this pattern ?![[A-z]{3,5}\_[\w\d]
	# (?![[A-z]{3,5}\_[\w\d])
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

	return regexResults	

# Takes in list of slides
# Returns naturally sorted list of slides
def sortFiles(fileList):
	# The very, very lazy mans sort:
	# The list has them slide10, slide 11, ..., slide 1
	# We want a more natural sort to maintain slide order
	# We achieve this by parsing each slide name down to its #, then sorting before adding context back
	# There are far better ways to do it, but this is easy 
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