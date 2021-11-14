# The text is to be stored in this string
x = ''


def storeString(inString):
    global x
    x = inString
    testString = inString
    inString = ""
    length = len(testString)
    i = 0
 
    
    while i < length:
        if testString[i] != ",":
            inString += testString[i]
        else:
            inString += "."
        i += 1

    try:
        inString = float(inString)
        if inString > 0:
            return inString
        else:
            return "0"
    except:
        return "0"

    
