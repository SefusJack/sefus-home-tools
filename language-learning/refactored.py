import urllib.request, json
import requests 
import openpyxl
from openpyxl import Workbook
from bs4 import BeautifulSoup
import re

kanjistash = list()
kanjidata = list()

pound="%23"
n5=pound+"jlpt-n5"
n4=pound+"jlpt-n4"
n3=pound+"jlpt-n3"
n2=pound+"jlpt-n2"
n1=pound+"jlpt-n1"
definition=""

english = ["Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", 
            "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ",
            "Ｘ", "Ｙ", "Ｚ", "０", "１", "２", "３", "４", "５", "６", "７", "８", "９"]

try:
    wb = openpyxl.load_workbook("sample.xlsx")
    ws = wb["Sheet"]
    ws2 = wb["Kanji"]
except:
    wb = Workbook()
    ws = wb.active
    ws2 = wb.create_sheet("Kanji")
    ws.append(["Word", "Meaning", "Tags", "Parts of Speech", "Kanji 1", "Meaning 1", "Kunyomi 1", "Onyomi 1", "Kanji 2", "Meaning 2", "Kunyomi 2", "Onyomi 2", "Kanji 3", "Meaning 3", "Kunyomi 3", "Onyomi 3", "Kanji 4", "Meaning 4", "Kunyomi 4", "Onyomi 4" ])
    ws2.append(["Kanji", "Meaning", "Frequency", "Kunyomi", "Onyomi"])

#Returns a 2D list of all the Rows data from an excel worksheet
def getDataFromAllExcelRows(worksheet, header):
    rowsdata = list()
    for i in range(worksheet.min_row+header, worksheet.max_row):
        rowsdata.append(getDataFromExcelRow(worksheet, i))
    return rowsdata

#Returns a list of the data from an inputted row and excel worksheet
def getDataFromExcelRow(worksheet, row):
    rowdata = list()
    for col in worksheet[row]:
        rowdata.append(col.value)
    return rowdata

#Returns a 2D list of all the Columns data from an excel worksheet
def getDataFromAllExcelColumns(worksheet, header):
    colsdata = list()
    for i in range(worksheet.min_column, worksheet.max_column):
        colsdata.append(getDataFromExcelColumn(worksheet, header, i-1))
    return colsdata

#Returns a list of the data from an inputted column and excel worksheet
def getDataFromExcelColumn(worksheet, header, column):
    coldata = list()
    for row in worksheet.iter_rows(worksheet.min_row+header,worksheet.max_row):
        coldata.append(row[column].value)
    return coldata

#Returns the data from a single cell from a worksheet
def getDataFromCell(worksheet, row, column):
    return worksheet.cell(row, column).value

#Returns True if the inputted character is a kanji, Returns false otherwise
def isKanji(character):
    return re.match(r'([^一-龯])', character) is None

#Returns a list of kanji found in a word
def getKanjiFromWord(word):
    temp = list()
    if word is not None:
        for character in word:
            if isKanji(character):
                temp.append(character)
    return temp

#Returns a list of kanji from a list of words
def getKanjiListFromWords(words):
    kanjistemp = list()
    for word in words:
        kanjis = getKanjiFromWord(word)
        for kanji in kanjis:
            if kanji not in kanjistemp:
                kanjistemp.append(kanji)
    return kanjistemp

#Returns a list of kanji from a list of words (including only kanji) that are currently in the inputted excel worksheet, header number, and column to search in
def getKanjiListFromExcelWorkSheet(worksheet, header, column):
    return getKanjiListFromWords(getDataFromExcelColumn(worksheet, header, column))

#Returns true if the Kanji inputted is found in the Kanji stash, returns false otherwise
def inKanjiStash(kanji):
    return kanji in kanjistash

#Adds a List of Kanji to the Kanji Stash
def addKanjiListToStash(kanjis):
    for kanji in kanjis:
        if not inKanjiStash(kanji):
            kanjistash.append(kanji)

#Adds a Kanji to the Kanji stash
def addKanjiToStash(kanji):
    if not inKanjiStash(kanji):
        kanjistash.append(kanji)

#Adds all of the Kanji from the Kanji Stash To Excel. Kanji already in the Excel WorkSheet will be ignored
def addKanjiStashToExcel(worksheet, header, column):
    for kanji in kanjistash:
        temp = getKanjiListFromExcelWorkSheet(worksheet, header, column)
        if kanji not in temp:
            worksheet.append([kanji])

def getKanjiDataFromKanjiStash():
    for kanji in kanjistash:
        kanjidata.append(kanjiSearch(kanji))

###########################################################################################################
def getWord(data):
    if 'word' in data['japanese'][0]:
        return data['japanese'][0]['word']
    else:
        return getReading(data)

def getReading(data):
    return data['japanese'][0]['reading']

def getMeaning(data, limit, multiwordf):
    temp = ""
    index = 1
    length = len(data['senses'])
    if data['senses'][length-1]['parts_of_speech']:
        if data['senses'][length-1]['parts_of_speech'][0] == 'Wikipedia definition':
            length -= 1
    for i in data['senses']:
            if index < length and index < limit:
                temp = temp + '"' + str(index) + ". " + ", ".join(i['english_definitions']).replace('"','') + '"' + " &CHAR(10)& "
            else:
                if not multiwordf:
                    temp = temp + '"' + str(index) + ". " + ", ".join(i['english_definitions']).replace('"','') + '"'
                    return "= " + temp
                else:
                    return ", ".join(i['english_definitions']).replace('"','')
            index += 1

def getJLPT(data):
    if data['jlpt']:
        return list(reversed(sorted(data['jlpt'])))[0]
    else:
        return ""

def getCommonality(data):
    if data['is_common'] != None:
        if data['is_common'] == True:
            return "common"
        else:
            return "uncommon"
    else:
        return ""

def getPartOfSpeech(data):
    temp = []
    for i in data['senses'][0]['parts_of_speech']:
        if i == "Suru verb" or i == "Godan verb with 'ru' ending" or i == "Intransitive verb" or i == "Ichidan verb" or i == "Transitive verb" or i == "Godan verb with 'mu' ending" or i == "Godan verb with 'su' ending" or i == "Godan verb with 'u' ending" or i == "Godan verb with 'ku' ending" or i == "Kuru verb - special class" or i == "Suru verb - included" or i == "Godan verb with 'bu' ending" or i == "Godan verb with 'gu' ending" or i == "Noun or verb acting prenominally" or i == "Godan verb with 'nu' ending" or i == "Irregular nu verb" or i == "Godan verb - Iku/Yuku special class" or i == "Godan verb with 'tsu' ending" or i == "Godan verb with 'ru' ending (irregular verb)" or i == "Ichidan verb - kureru special class" or i == "Godan verb - -aru special class" or i == "Auxiliary verb":
            temp.append("Verb")
        elif i == "Adverb (fukushi)" or i == "Adverb taking the 'to' particle":
            temp.append("Adverb")
        elif i == "Na-adjective (keiyodoshi)" or i == "I-adjective (keiyoushi)" or i == "Pre-noun adjectival (rentaishi)":
            temp.append("Adjective")
        elif i == "Noun which may take the genitive case particle 'no'" or i == "Noun, used as a prefix" or i == "Noun, used as a suffix":
            temp.append("Noun")
        elif i == "Expressions (phrases, clauses, etc.)":
            temp.append("Expressions")
        else:
            temp.append(i)
        
    temp = list(dict.fromkeys(temp))
    return  " ".join(temp)

def getKanji(word):
    temp = list()
    for i in list(word):
        if i not in hiragana and i not in katakana and i not in halfwidth and i not in english:
            temp.append(i)
            if i not in kanji:
                kanji.append(i)
                kanjidata.append(kanjiSearch(i))
    return temp

def isEnglish(word):
    for i in list(word):
        if i in english:
            return True
    return False

def isOnlyKanji(word):
    for i in list(word):
        if not isKanji(i):
            return False
    return True

def getFurigana(word, reading):
    if not isEnglish(word):
        wordcharacters = list(word)
        readingcharacters = list(reading)
        readingcharlist = list()
        temp = list()
        furiword = ""
        index = 0
        for char in readingcharacters:
            if char not in wordcharacters:
                temp.append(char)
                if index == len(readingcharacters)-1:
                    readingcharlist.append("".join(temp))
                    temp = list()
                elif readingcharacters[index+1] in wordcharacters:
                        readingcharlist.append("".join(temp))
                        temp = list()
            else:
                temp.append(char)
                readingcharlist.append(char)
                temp = list()
            index += 1

        index = 0
        furiPos = 0
        for char in wordcharacters:
            if isKanji(char):
                temp.append(char)
                if index == len(wordcharacters)-1:
                    print(word)
                    print(reading)
                    print(char)
                    furiword = furiword + "".join(temp) + "[" + readingcharlist[furiPos] + "]"
                    furiPos += 1
                    temp = list()
                elif not isKanji(wordcharacters[index+1]):
                        furiword = furiword + "".join(temp) + "[" + readingcharlist[furiPos] + "]"
                        furiPos += 1
                        temp = list()
            else:
                temp.append(char)
                furiword = furiword + "".join(temp)
                furiPos += 1
                temp = list()
            index += 1

        return furiword
    else:
        return reading

def readPage(start, limit, search):
    page = start
    if limit == 0:
        limit = sys.maxint
    data = list()
    while(page <= limit):
        print("Page Number:" + str(page))
        with urllib.request.urlopen("https://jisho.org/api/v1/search/words?keyword=" + search + "&page=" + str(page)) as site:
            result = json.loads(site.read().decode())
        if len(result['data']) == 0:
            return data
        else:
            data.append(result)
        page += 1
    return data

def wordSearchToExcel(start, limit, search):
    data = readPage(start, limit, search)
    for i in data:
        for j in i['data']:
            word = getWord(j)
            reading = getReading(j)
            meaning = getMeaning(j, 3, False)
            kanjiused =  getKanjiFromWord(word)
            furigana = getFurigana(word, reading)
            partsofspeech = getPartOfSpeech(j)
            tags = [getJLPT(j),getCommonality(j), partsofspeech]
            ws.append([furigana, meaning, " ".join(tags), partsofspeech.replace(' ',', ')])

def getListOfKanjiInKanjiData():
    kanjis = list()
    for i in kanjidata:
        kanjis.append(i[0])
    return kanjis

def getIndexOfKanjiInKanjiData(kanji):
    return getListOfKanjiInKanjiData().index(kanji)

def kanjiSearch(kanji):
    if kanji not in getListOfKanjiInKanjiData():
        url = requests.get("https://jisho.org/search/" + str(kanji) + "%20%23kanji")
        soup = BeautifulSoup(url.content, 'html.parser')

        frequency = ""
        meaning = ""
        kunyomi = ""
        onyomi = ""
        grade = ""
        jlpt = ""
        #Newspaper Frequency
        if soup.find_all("div", {"class": "frequency"}):
            frequency = soup.find_all("div", {"class": "frequency"})[0].text.strip()
        #Kanji Meaning
        if soup.find_all("div", {"class": "kanji-details__main-meanings"}):
            meaning = soup.find_all("div", {"class": "kanji-details__main-meanings"})[0].text.strip()
        #Kunyomi
        try:
            if soup.find_all("dd", {"class": "kanji-details__main-readings-list"}):
                kunyomi = soup.find_all("dd", {"class": "kanji-details__main-readings-list"})[0].text.strip()
        except:
            pass
        #Onyomi
        try:
            if soup.find_all("dd", {"class": "kanji-details__main-readings-list"}):
                onyomi = (soup.find_all("dd", {"class": "kanji-details__main-readings-list"})[1].text.strip()).translate(kat2hir)
        except:
            pass
        
        #Grade
        try:
            if soup.find_all("div", {"class": "grade"}):
                grade = soup.find_all("div", {"class": "grade"})[0].text.strip()
                grade = grade[grade.find('grade'):].replace(' ','-')
        except:
            pass

        #JLPT
        try:
            if soup.find_all("div", {"class": "jlpt"}):
                jlpt = soup.find_all("div", {"class": "jlpt"})[0].text.strip()
                jlpt = "jlpt-" + jlpt[11:].lower()
        except:
            pass
        tags = jlpt + " " + grade
        words = readPage(1, 1, "*" + urllib.parse.quote(str(kanji) + "*"))
        definitions = ""
        index = 0
        for j in words:
                for k in j['data']:
                    if index < 5:
                        defword = getWord(k)
                        defreading = getReading(k)
                        defmeaning = getMeaning(k, 1, True)
                        deffurigana = getFurigana(defword, defreading)
                        definitions = definitions + '"' + deffurigana + " " + defmeaning + '"' + " &CHAR(10)& "
                        index += 1
        definitions = "= " + definitions[:-12]
        info = [kanji, meaning, frequency, kunyomi, onyomi, definitions, tags]
        kanjidata.append(info)
    else:
        info = kanjidata[getIndexOfKanjiInKanjiData(kanji)]
    return info
    

def addKanjiDataToWords(worksheet, header, column):
    addKanjiListToStash(getKanjiListFromExcelWorkSheet(worksheet, header, column))
    words = getDataFromExcelColumn(worksheet, header, column)
    currentrow = header+1
    for kanjis in words:
        temp = getKanjiFromWord(kanjis)
        currentcol = 5
        for kanji in temp:
            data = kanjidata[getIndexOfKanjiInKanjiData(kanji)]
            worksheet.cell(currentrow, currentcol).value = kanji
            currentcol += 1
            worksheet.cell(currentrow, currentcol).value = data[1]
            currentcol += 1
            worksheet.cell(currentrow, currentcol).value = data[3]
            currentcol += 1
            worksheet.cell(currentrow, currentcol).value = data[4]
            currentcol += 1
        currentrow += 1


def kanjiSearchToExcel(kanji):
    for i in kanji:
        if i not in getListOfKanjiInKanjiData():
            info = kanjiSearch(i)
            ws2.append(info)

def getOnyomi(kanji):
        return kanjiSearch(kanji)[4]

def getKunyomi(kanji):
        return kanjiSearch(kanji)[3]


#Startup
kanjistash = getKanjiListFromExcelWorkSheet(ws, 1, 0)
kanjidata = getDataFromAllExcelRows(ws2, 1)
#Start Here
#wordSearchToExcel(1, 100, n4)
addKanjiListToStash(getKanjiListFromExcelWorkSheet(ws, 1, 0))
kanjiSearchToExcel(kanjistash)
addKanjiDataToWords(ws, 1, 0)

#End
wb.save("sample.xlsx")