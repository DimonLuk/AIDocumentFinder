"""
This module is the collection of functions.
Using them you are able to
1) Scan target file(.doc or .pdf) and find which words are the most common for it
2) Donwload .doc or .pdf files from google
3) Scan downloaded files and decide if they are usable
To know imports see every function
TODO Deal with pdf
TODO make analytics on words
TODO: create neural network which will make decision for you
"""
debug = True
import win32com.client
def getTextFromWordDocument(path, fileName):
    """
    path to file in format: "Disk:\Path\To\File\"
    fileName in format: "Document.doc(x)"
    Returnes text from Document.doc(x)
    Imports win32com.client
    """
    import win32com.client
    app = win32com.client.Dispatch("Word.Application")
    pathToFile = r"%s\%s" % (path, fileName)
    try:
        return app.Documents.Open(pathToFile).Content.Text
    except win32com.client.pywintypes.com_error:
        return ""
    except AttributeError:
        return ""




    

def getTextFromPdfDocument(path, fileName): #TODO
    pass


def createTupleOfWords(text):
    """
    text which is returned by getTextFromWordDocument or  getTextFromPdfDocument
    Returnes array of words which is formated to lower case
    Imports re
    """
    import re
    text = text.lower()
    text = re.findall("\w+", text)
    return tuple(text)

def countEveryWord(words,*, filter=[], wordLengthMoreThan=4):
    """
    words - array of words in lower case.
    filter - optional array of words to be found and counted
    wordLengthMoreThan - optional. it's the minimum length of the word
    Returns tuple which contains Counter({"word":int(homMuchAppearsInTheArray)}) and total sum of each word in the counter
    filter has more priority than wordLengthMoreThan
    Imports: collections.Counter
    """
    from collections import Counter
    c = Counter(words)
    totalNum = sum(c.values())
    found = []
    #Deletes trash words
    if (not wordLengthMoreThan) and (not filter):
        del c["та"]
        del c["на"]
        del c["таеп"]
        del c["з"]
        del c["в"]
        del c["і"]
        del c["що"]
        del c["зв"]
        del c["їх"]
        del c["ня"]
        del c["для"]
        del c["gc"]
        del c["of"]
        del c["the"]
        del c["при"]
        del c["за"]
        del c["end"]
        del c["and"]
        for i in range(1000):
            tmp = "%s" % i
            del c[tmp]
        alphabet = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЬЫЪЭЮЯАБВГДЕЄЖЗИІЇЙКЛМНОПРСТУФХЦЧШЩЬЮЯABCDEFGHIJKLMNOPQRSTUVWXYZ"
        lalphabet = alphabet.lower()
        for i in alphabet:
            del c[i]
        for i in lalphabet:
            del c[i]
    elif filter:
        for item in c:
            for i in filter:
                if i in item:
                    found.append(i)
    elif wordLengthMoreThan:
        for item in c:
            if len(item) >= wordLengthMoreThan:
                found.append(item)
    c = Counter(found)
    #Return the result
    return tuple((c,sum(c.values()), totalNum))



def getResultOfCounting(infoAboutWords, accuracy, infoJson):
    """
    infoAboutWords - tuple. The first element - Counter({"word":int(homMuchAppearsInTheArray)}), the second is total sum of each word in the counter
    accuracy - how much of the most popular words have to appear in the file
    fileName - the name of the file where words are counted
    file - is already opened file where to write info about files
    Writes info about each file with name "fileName" in opened file "file".
    Looks like:
    ---------------------------------------------
    fileName
    total sum of each word in the counter
    word:howMuchItAppears
    ---------------------------------------------
    Returns: nothing
    Imports: nothing
    """
    if infoAboutWords[1] != 0:
        infoJson["totalNum"] = infoAboutWords[2]
        infoJson["totalPopNum"] = infoAboutWords[1]
        words = {}
        for word in infoAboutWords[0].most_common(accuracy):
            try:
                words[word[0]] = word[1]
            except UnicodeEncodeError:
                pass
        infoJson["words"] = words
    return infoJson


def downloadDocuments(pathToSave, extension="doc", *, fromInternet=False, link="", pathToFiles=""):
    """
    Two ways of using:
    1) Save htmls automaticaly(Dont work)
    2) Save htmls manualy
    pathToSave - where to save files in format "Disk:\Path\To\Save"
    extension - defines which type is used("doc" or "pdf"). "doc" is the same as "docx" but "doc" is more general and has to be used
    fromInternet - defines if the app should use manualy saved htmls or get them form link
    link - defines link which is used to get pages from the Google directly
    pathToFiles - defines where local htmls saved

    Downloads documents to pathToSave
    Returns: Nothing
    Imports: urllib.request, bs4(BeautifulSoup)
    """
    import urllib.request
    from bs4 import BeautifulSoup
    import re
    import json
    pages=[]
    schema = r".*\.%s" % extension
    schema = re.compile(schema)
    infos = []
    #Get pages from the Internet or from local files
    if fromInternet:
        pages = getPagesFromTheInternet(link)
    else:
        pages = getPagesFromFiles(pathToFiles)
    
    #Open file, where source links have to be located 
    infoJson = r"%s\%s" % (pathToSave, "info.json")
    #Then iterate through each page
    for page in pages:
        soup = BeautifulSoup(page, "lxml")
        hrefContainer =  soup.find_all("h3", {"class":["r"]})#h3 contains <a> with the href to download files
        for item in hrefContainer:
            info = {"name":"", "totalNum":0, "totalPopNum":0, "words":{}, "link":""}
            if fromInternet:
                fullLink = item.find("a").get("href")[7:]#if page is from the Internet it looks like /?url?=htttp(s)://...
            else:
                fullLink = item.find("a").get("href")#if page is from file its href looks like http(s)://...
            if debug:
                print(fullLink)
            #print looks like ---------------------\nfullLink\n--------------------
            info["link"] = fullLink
            #if page is from th Internet it can contain some junk after download link, regexpr deletes this junk
            for link in schema.findall(fullLink):
                if link:#Throw empty links
                    fileName = ""
                    parts = link.split("/")
                    #Create normal name
                    for part in parts:
                        if (".%sx" % extension) in part:
                            if len(part) > 20:
                                fileName = part[len(part)-10:]+"."+extension
                            fileName = part
                        elif (".%s" % extension) in part:
                            if len(part) > 20:
                                fileName = part[len(part)-10:]+"."+extension
                            fileName = part
                            info["name"] = fileName
                    fullPath = r"%s\%s" % (pathToSave, fileName)
                    try:
                        with open(fullPath, "wb") as file:
                            try:
                                with urllib.request.urlopen(link) as download:
                                    file.write(download.read())
                            except:
                                pass
                    except FileNotFoundError:
                        pass
                    except OSError:
                        pass
            infos.append(info)
    with open(infoJson, "w") as file:
        json.dump(infos, file, ensure_ascii=False, indent=1)

def getPagesFromFiles(pathToFiles):
    """
    pathToFiles (in format: "Disk:\Path\To\Files") defines where manually saved htmls are stored
    iterates over each file in the pathToFiles dir and create array of pages
    Returns: array of pages
    Imports: os
    """
    import os
    pages = []
    i = 0
    for dir, subdirs, files in os.walk(pathToFiles):
        for file in files:
            absPath = "%s\%s" % (pathToFiles, file)
            with open(absPath, "rb") as doc:
                pages.append(doc.read().decode("utf-8"))#Has to decode as utf-8 because these are google's pages
                
    return pages
def getPagesFromTheInternet(link):#TODO
    pass
def countWordsInFiles(pathToFiles,*, extension="doc"):
    """
    pathToFiles in format "Drive:\Path\to\files" defines where downloaded documents.extension are stored
    extension defines extension of the file: "doc", "pdf". "doc" and "docx" are the same but "doc" has to be used
    Opens file pathToFile\info.txt where info about each file will be written using function getResultOfCounting
    Returns: nothing
    Imports: os
    """
    import os
    import json
    info = []
    with open("%s\info.json" % pathToFiles, "r") as file:
        info = json.load(file)
    i = 0
    """for dir, subdirs, files in os.walk(pathToFiles):
        numOfFiles = len(files)
        if debug:
            print(numOfFiles)
        for file in files:
            if i == numOfFiles:
                break
            if extension == "doc":#Check if it's not something else
                if (".doc" in file):
                    getResultOfCounting(countEveryWord(createTupleOfWords(getTextFromWordDocument(pathToFiles, file)), filter=["синх", "диагнос", "двиг", "метод", "сравн", "парам"]), 20, file, info)
            if debug:
                text = "Proceed: %.d" % ((i/numOfFiles)*100)
                text += "%"
                print(text)
                print(i)
                print(file)
            i += 1"""
    for file in info:
        file = getResultOfCounting(countEveryWord(createTupleOfWords(getTextFromWordDocument(pathToFiles, file["name"]))), 20, file)
    with open("%s\info.json" % pathToFiles, "w") as file:
        json.dump(info, file, ensure_ascii=False, indent=1)



def delteTrash(path):
    with open("%s\info.txt" % path, "r") as file:
        print(file.readline())
#----------------------Analysing--------------------------
def getInfo(path, extension="doc"):
    """
    path in format "Drive:\path" defines where file with info about documents is stored
    accuracy defines how many of the most popular words will be calculated
    Reads files and finds how many times the most popular words from files appers in the info.txt
    Returns: nothing
    Imports: re, collections
    """
    import re
    import collections
    pattern1 = re.compile(r"^\d+$")
    pattern2 = re.compile((r"\w+\.%s") % extension)
    info = []
    tmp = []
    counter = 0
    with open("%s\info.txt" % path, "r") as file:
        while file.readline():
            text = file.readline()
            if pattern1.findall(text):
                tmp.append(int(pattern1.findall(text)[0]))
                info.append(tuple(tmp))
                tmp = []
            if "---------------------------------------------" in text:
                counter += 1
            if pattern2.findall(text):
                tmp.append(pattern2.findall(text)[0])
    print(counter)
    print(info)
    return info
def deleteJunkyFiles(path, extension="doc"):
    import os
    with open(("%s\info.txt" % path), "r") as infoTxt:
        for dir, subdirs, files in os.walk(path):
            for file in files:
                if extension in file:
                    if not (file in infoTxt.read()):
                        os.remove("%s\%s" % (path, file))
def plotNumsOfWords(infoToPlot):
    import matplotlib.pyplot as plt
    x = infoToPlot
    y = infoToPlot
    plt.scatter(x, y, marker="^")
    plt.show()

    
if __name__ == "__main__":
    if debug:
        #downloadDocuments(r"D:\Projects\AIIDocumentFinder\Alpha\AIDocumentFinder\docs", pathToFiles="D:\Projects\AIIDocumentFinder\Alpha\AIDocumentFinder\htmls")
        countWordsInFiles("D:\Projects\AIIDocumentFinder\Alpha\AIDocumentFinder\docs")
        #delteTrash("D:\Projects\AIIDocumentFinder\Alpha\AIDocumentFinder\docs")
        #getInfo("D:\Projects\AIIDocumentFinder\Alpha\AIDocumentFinder\docs")
        #deleteJunkyFiles("D:\Projects\AIIDocumentFinder\Alpha\AIDocumentFinder\docs")