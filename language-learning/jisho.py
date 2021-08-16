import urllib.request, json 
from openpyxl import Workbook


wb = Workbook()

#grab the active worksheet
ws = wb.active

pound="%23"
n5=pound+"jlpt-n5"
n4=pound+"jlpt-n4"
n3=pound+"jlpt-n3"
n2=pound+"jlpt-n2"
n1=pound+"jlpt-n1"
definition=""

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
other = ["ー", "ィ", "ォ", "ッ", "ャ", "ュ", "っ", "ゃ", "々", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", 
            "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ",
            "Ｘ", "Ｙ", "Ｚ", "０", "１", "２", "３", "４", "５", "６", "７", "８", "９"]

#Data can be assigned directly to cells
ws.append(["Kanji", "Reading", "Meaning", "Tags"])

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
        if i not in hiragana and i not in katakana and i not in other:
            temp.append(i)
    return ", ".join(temp)

def hasOnlyKanji(word):
    temp = list()
    for i in list(word):
        if i not in hiragana and not i in katakana and i not in other:
            return true
    return false

def hasOnlySimpleKana(word):
    temp = list()
    for i in list(word):
        if i not in hiragana and not i in katakana and i not in other:
            return true
    return false

def getFurigana(word, reading):
    temp = list()
    for i in list(word):
        if i not in hiragana and not i in katakana and i not in other:
            return 
    return true

def readPage(page):
    print("Page Number:" + str(page))
    with urllib.request.urlopen("https://jisho.org/api/v1/search/words?keyword=%23jlpt-n5&page=" + str(page)) as url:
        data = json.loads(url.read().decode())
        if len(data['data']) != 0:
            for i in data['data']:
                word = getWord(i)
                reading = getReading(i)
                meaning = getMeaning(i)
                kanji = getKanji(word)
                tags = [getJLPT(i),getCommonality(i), getPartOfSpeech(i)]
                ws.append([word, reading, meaning, " ".join(tags), kanji])
            readPage(page + 1)

readPage(1)
#Rows can also be appended
#ws.append(["sefus", test, 3])

#Save the file
wb.save("sample.xlsx")