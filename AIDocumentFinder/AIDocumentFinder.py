"""
This odule if the collection of functions.
Using them you are able to
1) Scan target file(.doc or .pdf) and find which words are the most common for it
2) Donwload .doc or .pdf files from google
3) Scan downloaded files and decide if they are usable
imports: win32com.client
TODO Deal with pdf
TODO: create neural network which will make decision for you
"""
debug = True
import win32com.client
def getTextFromWordDocument(path, fileName, *, debugging=False):
    """
    Using: path where will be created temporary files, filename with .doc or pdf
    Returns: text from original file
    imports win32com.client
    """
    import win32com.client
    app = win32com.client.Dispatch("Word.Application")
    pathToFile = r"%s\%s" % (path, fileName)
    return app.Documents.Open(pathToFile).Content.Text



    

def getTextFromPdfDocument(path, fileName, *, debugging=False): #TODO
    """
    Using: path where will be created temporary files, filename with .doc or pdf
    Returns: text from original file
    """
    from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer.converter import TextConverter
    from pdfminer.layout import LAParams
    from pdfminer.pdfpage import PDFPage
    from io import StringIO
    
    pathToFile = r"%s\%s" % (path, fileName)
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(pathToFile, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text[:-1]


def createArrayOfWords(text, *, debuggin=False):
    """
    Input: text,
    imports: re
    """
    import re
    text = text.lower()
    text = re.findall("\w+", text)
    return text

def countEveryWord(words):
    """
    Using: counts how many times each word appears in the array of words
    than delete words witihout any information like "the", "and" and so on
    """
    from collections import Counter
    c = Counter(words)
    #Deletes trash words
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
    #Return the result
    return c



def getResultOfCounting(countedWords, accuracy, path, fileName, *, write=False):
    """
    Using with write=False - returns info about most common words which number is set with accuracy
    Using with write=True - writes to file info about most common words which number is set with accuracy
    """
    if write:
        pathToFile = r"%s\readyToExplore.txt" % path
        with open(pathToFile, "w") as file:
            text = "---------------------------------------------------\n%s\n" % fileName
            file.write(text)
            for i in countedWords.most_common(accuracy):
                file.write("%s:%s\n" % i)
            file.write("---------------------------------------------------\n")
    if not write:
        return countedWords.most_common(accuracy)


def downloadDocuments(pathToSave, *, fromInternet=False, link="", pathToFiles=""):
    import urllib.request
    from bs4 import BeautifulSoup
    import re
    import shutil
    pages=[]
    if fromInternet:
        pages = getPagesFromTheInternet(link)
    else:
        pages = getPagesFromFiles(pathToFiles)
    fileWithSources = "%s\%s" % (pathToSave, "sources.txt")
    fileWithSources = open(fileWithSources, "a")
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.78 Safari/537.36 OPR/47.0.2631.55"}
    for page in pages:
        soup = BeautifulSoup(page, "lxml")
        hrefContainer =  soup.find_all("h3", {"class":["r"]})
        for item in hrefContainer:
            if fromInternet:
                fullLink = item.find("a").get("href")[7:]
            else:
                fullLink = item.find("a").get("href")
            print(fullLink)
            fileWithSources.write("-----------------------------------------\n")
            fileWithSources.write(fullLink)
            fileWithSources.write("\n")
            for link in re.findall(r".*\.doc", fullLink):
                if link:
                    fileName = ""
                    parts = link.split("/")
                    for part in parts:
                        if ".doc" in part:
                            fileName = part
                    fullPath = r"%s\%s" % (pathToSave, fileName)
                    try:
                        with open(fullPath, "wb") as file:
                            with urllib.request.urlopen(link) as download:
                                file.write(download.read())
                        if fromInternet:
                            fileWithSources.write(link)
                            fileWithSources.write("\n")
                    except FileNotFoundError:
                        pass
            fileWithSources.write("-----------------------------------------\n")
    fileWithSources.close()

def getPagesFromFiles(pathToFiles):
    import os
    pages = []
    try:
        for dir, subdirs, files in os.walk(pathToFiles):
            for file in files:
                absPath = "%s\%s" % (pathToFiles, file)
                with open(absPath, "rb") as doc:
                    pages.append(doc.read().decode("utf-8"))
    except StopIteration:
        pass
    return pages
def getPagesFromTheInternet(link):
    pass
    
if __name__ == "__main__":
    if debug:
        #print(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "Test.doc"))
        #print(getTextFromPdfDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "Test.pdf"))
        #print(createArrayOfWords(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "Test.doc")))
        #print(createArrayOfWords(getTextFromPdfDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "realtest1.pdf")))
        #print(createArrayOfWords(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "realtest1.doc")))
        #print(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "Test.doc"))
        #print(createArrayOfWords(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "Test.doc")))
        #print(countEveryWord(createArrayOfWords(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "realtest1.doc"))))
        #countEveryWord(createArrayOfWords(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "realtest1.doc")))
        #print(getResultOfCounting(countEveryWord(createArrayOfWords(getTextFromWordDocument("D:\Projects\AIIDocumentFinder\AIDocumentFinder", "realtest.doc"))), 10, "D:\Projects\AIIDocumentFinder\AIDocumentFinder", "realtest.doc", write=True))
        #downloadDpcumentsFromGoogle("D:\Projects\AIIDocumentFinder\AIDocumentFinder\links.txt", "D:\Projects\AIIDocumentFinder\AIDocumentFinder")
        downloadDocuments("D:\Projects\AIIDocumentFinder\AIDocumentFinder\docs", pathToFiles="D:\Projects\AIIDocumentFinder\AIDocumentFinder\htmls")