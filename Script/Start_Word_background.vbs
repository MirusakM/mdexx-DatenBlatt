Dim objWord         ' Object for Word application

  Set objWord = CreateObject("Word.Application")
  objWord.Visible = False

  objWord.Documents.Open("C:\SAPZSIDACtmp\Datenblatt.docx")

