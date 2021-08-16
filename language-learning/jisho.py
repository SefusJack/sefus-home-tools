import urllib.request, json
import requests 
from openpyxl import Workbook
from bs4 import BeautifulSoup

wb = Workbook()

#grab the active worksheet
ws = wb.active
ws2 = wb.create_sheet("Kanji")

pound="%23"
n5=pound+"jlpt-n5"
n4=pound+"jlpt-n4"
n3=pound+"jlpt-n3"
n2=pound+"jlpt-n2"
n1=pound+"jlpt-n1"
definition=""
kanji = list()

hiragana = ["あ", "い", "う", "え", "お", 
            "か", "き", "く", "け", "こ", 
            "さ", "し", "す", "せ", "そ", 
            "た", "ち", "つ", "て", "と", 
            "な", "に", "ぬ", "ね", "の", 
            "は", "ひ", "ふ", "へ", "ほ", 
            "ま", "み", "む", "め", "も",
            "や",       "ゆ",       "よ",
            "ら", "り", "る", "れ", "ろ",
            "わ",                   "を",
            "が", "ぎ", "ぐ", "げ", "ご",
            "ざ", "じ", "ず", "ぜ", "ぞ",
            "だ", "ぢ", "づ", "で", "ど",
            "ば", "び", "ぶ", "べ", "ぼ"
            "ぱ", "ぴ", "ぷ", "ぺ", "ぽ",
            "ん"]
katakana = ["ア", "イ", "ウ", "エ", "オ",
            "カ", "キ", "ク", "ケ", "コ",
            "ガ", "ギ", "グ", "ゲ", "ゴ",
            "サ", "シ", "ス", "セ", "ソ",
            "ザ", "ジ", "ズ", "ゼ", "ゾ",
            "タ", "チ", "ツ", "テ", "ト",
            "ダ", "ヂ", "ヅ", "デ", "ド",
            "ナ", "ニ", "ヌ", "ネ", "ノ",
            "ハ", "ヒ", "フ", "ヘ", "ホ",
            "バ", "ビ", "ブ", "ベ", "ボ",
            "パ", "ピ", "プ", "ペ", "ポ",
            "マ", "ミ", "ム", "メ", "モ",
            "ヤ",       "ユ",       "ヨ",
            "ラ", "リ", "ル", "レ", "ロ",
            "ワ",                   "ヲ",
            "ン"]
halfwidth = ["ー", "ィ", "ォ", "ッ", "ャ", "ュ", "っ", "ゃ", "々"]
english = ["Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", 
            "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ",
            "Ｘ", "Ｙ", "Ｚ", "０", "１", "２", "３", "４", "５", "６", "７", "８", "９"]

#Data can be assigned directly to cells
ws.append(["Word", "Meaning", "Tags", "Kanji"])
ws2.append(["Kanji", "Meaning", "Frequency", "Onyomi", "Kunyomi"])
def getWord(data):
    if 'word' in data['japanese'][0]:
        return data['japanese'][0]['word']
    else:
        return getReading(data)

def getReading(data):
    return data['japanese'][0]['reading']

def getMeaning(data):
    temp = ""
    index = 1
    length = len(data['senses'])
    if data['senses'][length-1]['parts_of_speech']:
        if data['senses'][length-1]['parts_of_speech'][0] == 'Wikipedia definition':
            length -= 1
    for i in data['senses']:
            if index < length:
                temp = temp + '"' + str(index) + ". " + ", ".join(i['english_definitions']).replace('"','') + '"' + " &CHAR(10)& "
            else:
                temp = temp + '"' + str(index) + ". " + ", ".join(i['english_definitions']).replace('"','') + '"'
                return "= " + temp
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
        temp.append(i.replace(" ", "_"))
    return  " ".join(temp)

def getKanji(word):
    temp = list()
    for i in list(word):
        if i not in hiragana and i not in katakana and i not in halfwidth and i not in english:
            temp.append(i)
            if i not in kanji:
                kanji.append(i)
    return ", ".join(temp)

def isKanji(char):
    if char not in hiragana and not char in katakana and char not in halfwidth and char not in english:
        return True
    return False

#def hasOnlySimpleKana(word):
#    temp = list()
#    for i in list(word):
#        if i not in hiragana and not i in katakana and i not in english:
#            return true
#    return false

def isEnglish(word):
    for i in list(word):
        if i in english:
            return True
    return False

def getFurigana(word, reading):
    wordcharacters = list(word)
    readingcharacters = list(reading)
    furichars = list()
    furiword = ""
    lastkanjipos = -1
    for i in readingcharacters:            
        if i not in wordcharacters:
            furichars.append(i)

    index = 0
    for j in wordcharacters:
        index += 1
        if isKanji(j):
            lastkanjipos = index
        if isEnglish(word):
            return reading
    
    if lastkanjipos > -1:
        return word[:lastkanjipos] + "[" + "".join(furichars) + "]" + word[lastkanjipos:]
    return word

def readPage(page):
    if(page < 10):
        print("Page Number:" + str(page))
        with urllib.request.urlopen("https://jisho.org/api/v1/search/words?keyword=%23jlpt-n5&page=" + str(page)) as url:
            data = json.loads(url.read().decode())
            if len(data['data']) != 0:
                for i in data['data']:
                    word = getWord(i)
                    reading = getReading(i)
                    meaning = getMeaning(i)
                    kanji = getKanji(word)
                    furigana = getFurigana(word, reading)
                    tags = [getJLPT(i),getCommonality(i), getPartOfSpeech(i)]
                    ws.append([furigana, meaning, " ".join(tags), kanji])
                readPage(page + 1)

readPage(1)
for i in kanji:
    url = requests.get("https://jisho.org/search/" + str(i) + "%20%23kanji")
    soup = BeautifulSoup(url.content, 'html.parser')
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
            onyomi = soup.find_all("dd", {"class": "kanji-details__main-readings-list"})[1].text.strip()
    except:
        pass
    ws2.append([i, meaning, frequency, kunyomi, onyomi])
#Rows can also be appended
#ws.append(["sefus", test, 3])

#Save the file
wb.save("sample.xlsx")