from shareplum import Site, Office365
from shareplum.site import Version
from threading import Thread
from multiprocessing.pool import ThreadPool
#returns a site object, needs valid username and password in order to work.
def getSite(siteurl: str, username: str, password: str):
    authurl = siteurl.partition("sharepoint.com/")
    authurl = authurl[0] + authurl[1]
    auth = Office365(share_point_site=authurl, username=username, password=password).GetCookies()
    return Site(site_url=siteurl, authcookie=auth, version=Version.v365)

#returns a sharepoint list from the site object.
def getList(site, listname: str):
    return site.List(listname, True)

#creates a new header
def createHeader(list, headerName: str, headerType: int): 
    list.create_field(headerName, headerType)

# checks to see if the sharepoint list has a specic header
def hasHeader(list, name: str):
    fields = list.fields
    for field in fields:
        if(field["DisplayName"] == name):
            return True
    return False
#Adds headers if they dont exist in the sharepoint list
def addHeadersConditional(excelSheet, list, headerType: int = 3):
     for col in excelSheet.columns:
        if(hasHeader(list=list, name=col) == False):
            list.create_field(title=col, field_type=headerType)
            print("Created Header: " + col)
#Adds data from excel sheet to sharepoint list, does NOT UPDATE the data
def addDataToList(list, excelSheet, threaded: bool = False, sortThreads: int = 1):
    data = sortToList(excelDoc=excelSheet, threads=sortThreads)
    fData = listSplitter(data, maxLength=100)
    if(threaded):
        thread_list = []
        for d in fData:
            thread = Thread(target=uploadNewData, args=(d,list))
            thread_list.append(thread)
            thread.start()
        for t in thread_list:
            t.join()   
    else:
        for d in fData:
            uploadNewData(list=list, data=d)
def uploadNewData(data:list, list):
    list.UpdateListItems(data, "New")
    print("Uploaded: " + str(len(data)) + " items.")
    
#sorts an exceldoc to a list format which can be uploaded to a sharepoint list then returns the list
#goTo: int variable which is the number of rows that will be read in an excel file. -1 means read all rows
def sortToList(excelDoc, threads: int = 1):
    goTo = len(excelDoc["Category"].tolist())
    chunksize = int(goTo/threads)
    docs = [excelDoc.iloc[x:x+chunksize] for x in range(0, goTo, chunksize)]
    results = [] 
    returnList = []
    pool = ThreadPool(processes=threads)
    for doc in docs:
        result = pool.apply_async(func=format, args=(doc, len(doc)))
        results.append(result)
    for r in results:
        returnList.append(r.get()[0])
    return returnList


# takes in a list then returns a list of the split lists. maxLength is the max length per split list
def listSplitter(list: list, maxLength: int = 100):
    masterList = [list[x:x+maxLength] for x in range(0, len(list), maxLength)]
    return masterList
    
def format(excelDoc, endIndex: int = 1):
    fnList = []
    i = 0
    while(i < endIndex):
        dic = {}
        for col in excelDoc.columns:
            dic[col] = excelDoc[col].tolist()[i]
        print("Row " + str(i) + " sorted.")
        fnList.append(dic)
        i+=1
    return fnList
    
        