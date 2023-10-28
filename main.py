import win32gui
import win32ui
import win32con
import win32api
from python_imagesearch.imagesearch import imagesearch,imagesearcharea
from PIL import Image
import time


def searchImage(imagePath):
    pos = imagesearch(imagePath)
    if pos[0] != -1:
        return(pos[0], pos[1])
    else:
        print("image not found")

def saveScreenShot(x,y,width,height,path):
    # grab a handle to the main desktop window
    hdesktop = win32gui.GetDesktopWindow()

    # create a device context
    desktop_dc = win32gui.GetWindowDC(hdesktop)
    img_dc = win32ui.CreateDCFromHandle(desktop_dc)

    # create a memory based device context
    mem_dc = img_dc.CreateCompatibleDC()

    # create a bitmap object
    screenshot = win32ui.CreateBitmap()
    screenshot.CreateCompatibleBitmap(img_dc, width, height)
    mem_dc.SelectObject(screenshot)


    # copy the screen into our memory device context
    mem_dc.BitBlt((0, 0), (width, height), img_dc, (x, y),win32con.SRCCOPY)

    # save the bitmap to a file
    screenshot.SaveBitmapFile(mem_dc, path)
    # free our objects
    mem_dc.DeleteDC()
    win32gui.DeleteObject(screenshot.GetHandle())

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

def searchForColor(guess):
    guess -= 1
    listGreen = []
    listYellow = []
    listRed = []
    allColors = []
    allColorsText = []
    allColorsSorted = []
    allColorsSortedText = []
    go = True
    saveScreenShot(0,0,1920,1080,"screen.bmp")
    im = Image.open(r"screen.bmp")
    for j in range(100,640):
        for i in range(1920):
            if im.getpixel((i,j)) == (106,170,100):
                if listGreen == []:
                    listGreen.append((i,j))
                    allColorsText.append("Green")
                    allColors.append((i,j))
                for k in listGreen:
                    if abs(i - k[0]) < 64 and abs(j - k[1]) < 64:
                        go = False
                if go == True:
                    listGreen.append((i,j))
                    allColorsText.append("Green")
                    allColors.append((i,j))
                go = True
            elif im.getpixel((i,j)) == (201,180,88):
                if listYellow == []:
                    listYellow.append((i,j))
                    allColorsText.append("Yellow")
                    allColors.append((i,j))
                for k in listYellow:
                    if abs(i - k[0]) < 64 and abs(j - k[1]) < 64:
                        go = False
                if go == True:
                    listYellow.append((i,j))
                    allColorsText.append("Yellow")
                    allColors.append((i,j))
                go = True
            elif im.getpixel((i,j)) == (120,124,126):
                if listRed == []:
                    listRed.append((i,j))
                    allColorsText.append("Red")
                    allColors.append((i,j))
                for k in listRed:
                    if abs(i - k[0]) < 64 and abs(j - k[1]) < 64:
                        go = False
                if go == True:
                    listRed.append((i,j))
                    allColorsText.append("Red")
                    allColors.append((i,j))
                    go = False
                go = True
    for i in range(0+(guess*5),5+(guess*5)):
        allColorsSortedText.append(allColorsText[i])
    
    return(allColorsSortedText)
    
def findAreas():
    letter_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N",
                        "O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    positions = []
    for letter in letter_list:
        pos = searchImage("WordleUnlimited/" + letter + ".png")
        positions.append(pos)
    return(positions)

def clickWord(word,guess,positions):
    letter_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N",
                        "O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    word = word.upper()
    for i in word:
        pos = positions[letter_list.index(i)]
        click(pos[0]-1900,pos[1]+20)
    guess += 1
    clickEnter()
    return(guess)

def clickEnter():
    pos = searchImage("WordleUnlimited/Enter.png")
    click(pos[0]-1900,pos[1]+20)

def clickPlayAgain():
    time.sleep(2)
    pos = searchImage("WordleUnlimited/playAgain.png")
    click(pos[0]-1900,pos[1]+20)

class Letter:
    def __init__(self,letter,color,position,null_positions,inWord):
        self.letter = letter
        self.color = color
        self.position = position
        self.null_positions = null_positions
        self.inWord = inWord

def setUp():
    letter_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N",
                        "O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    letters = []
    for i in letter_list:
        letters.append(Letter(i,"None",[],[],False))
    wordTree = []
    guess = 0
    with open('guesstree.txt') as f:
        for line in f.readlines():
            wordTree.append(line.split(","))
            wordTree[-1][-1] = wordTree[-1][-1].strip("\n")
    return(letters,wordTree,guess)

def getWords(wordTree,letters):
    go = True
    a = 0
    for letter in letters:
        if letter.color == "None":
            continue
        elif letter.color == "Green":
            for i in wordTree:
                for e,j in enumerate(i):
                    for positions in letter.position:
                        if letter.letter.lower() not in j[positions] and go == True:
                            wordTree[wordTree.index(i)].remove(j)
                            a += 1
                            go = False
                    for position in letter.null_positions:
                        if letter.letter.lower() in j[position] and go == True:
                            wordTree[wordTree.index(i)].remove(j)
                            a += 1
                            go = False
                    go = True
        elif letter.color == "Red":
            for i in wordTree:
                for e,j in enumerate(i):
                    if letter.letter.lower() in j:
                        wordTree[wordTree.index(i)].remove(j)
                        a += 1
        elif letter.color == "Yellow":
            for i in wordTree:
                for e,j in enumerate(i):
                    if letter.letter.lower() not in j and go == True:
                        wordTree[wordTree.index(i)].remove(j)
                        a += 1
                        go = False
                        break
                    for position in letter.null_positions:
                        if letter.letter.lower() in j[position] and go == True:
                            wordTree[wordTree.index(i)].remove(j)
                            a += 1
                            go = False
                            break
                go = True
        go = True
    return(wordTree,a)

def getBestWord(wordTree,wordsGuessed):
    guess_dict = {}
    for i in wordTree:
        for j in i:
            if j not in wordsGuessed:
                if j in guess_dict:
                    guess_dict[j] += 1
                else:
                    guess_dict[j] = 1
    word = max(guess_dict, key=guess_dict.get)
    print(guess_dict)
    print(word,guess_dict[word])
    return(word)

def setLetterRed(letter,position,letters):
    if letters[letter].inWord != True:
        letters[letter].color = "Red"
    letters[letter].null_positions.append(position)
    return(letters)

def setLetterYellow(letter,position,letters):
    if letters[letter].inWord != True:
        letters[letter].color = "Yellow"
    letters[letter].null_positions.append(position)
    letters[letter].inWord = True
    return(letters)
    
def setLetterGreen(letter,position,letters):
    letters[letter].color = "Green"
    letters[letter].position.append(position)
    letters[letter].inWord = True
    return(letters) 
    
def main():
    letters,wordTree,guess = setUp()
    positions = findAreas()
    wordsGuessed = []
    a = 1
    b = 0
    while True:
        a = 1
        b = 0
        bestWord = getBestWord(wordTree,wordsGuessed)
        wordsGuessed.append(bestWord)
        print(wordsGuessed)
        guess = clickWord(bestWord,guess,positions)
        time.sleep(2)
        colors = searchForColor(guess)
        print(colors)
        if colors == ["Green","Green","Green","Green","Green"]:
            clickPlayAgain()
            time.sleep(2)
            letters,wordTree,guess = setUp()
            wordsGuessed = []
            a = 1
            b = 0
            continue
        for i in range(5):
            color = colors[i]
            if color == "Red":
                letters = setLetterRed(ord(bestWord[i])-97,i,letters)
            elif color == "Yellow":
                letters = setLetterYellow(ord(bestWord[i])-97,i,letters)
            elif color == "Green":
                letters = setLetterGreen(ord(bestWord[i])-97,i,letters)
        while a != 0:
            wordTree,a = getWords(wordTree,letters)
            b += 1

main()
    