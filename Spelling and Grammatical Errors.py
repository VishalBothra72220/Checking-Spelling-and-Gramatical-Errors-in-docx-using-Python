


import win32com.client as win32
Application=win32.gencache.EnsureDispatch('Word.Application')
Application.Visible=False
# Code for searching Spelling and Grammatical Errors
try:
    file=r"C:\Users\lenovo\Desktop\BotMantra\Task\highlights.docx"
    ActiveDocument=Application.Documents.Open(file)
    se=ActiveDocument.SpellingErrors.Count
    ge=ActiveDocument.GrammaticalErrors.Count
    print("Number of Spelling Errors:",se)
    print("Number of Grammatical Errors:",ge)
    print("Spelling errors are:-")
    for i in range(se):
        print(str(i+1)+")",ActiveDocument.SpellingErrors(i+1))
    print("Grammatical errors are:-")
    for i in range(ge):
        print(str(i+1)+")",ActiveDocument.GrammaticalErrors(i+1))
    #ActiveDocument.Close(True)
except:
    Application.Quit()

