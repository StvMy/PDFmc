import comtypes.client as comt

word = comt.CreateObject('Word.Application')

doc = word.Documents.Open("D:\\CODE\\1 test.docx")
doc.SaveAs("D:/CODE/output.pdf", FileFormat=17)  # 17 = PDF format
doc.Close()
word.Quit()

