'字段\手动替换字段所有内容(AllReplace): 

Option Public

Sub Initialize
	Dim ws As New NotesUiWorkspace
	Dim uidoc As notesuidocument
	Set uidoc=ws.currentdocument
	Dim ss As New NotesSession
	Dim db As NotesDatabase
	Dim Docc As NotesDocumentCollection
	Dim doc As NotesDocument, doc2 As NotesDocument
	Dim sField As String, sSize As String 
	Dim sValue As String
	Dim sTypeName As String 
	Dim i As Integer, k As Integer, y() As Variant, z As Variant
	Dim numFlag As Integer
	Set db = ss.currentdatabase
	Set Docc = db.unprocesseddocuments
	
	' 在功能表中是否有執行權限
	Dim level As Integer
	level = db.CurrentAccessLevel
	If level < ACLLEVEL_DESIGNER Then
		Messagebox "您無權執行此程式!",0,"警告"
		Exit Sub
	End If
	
	If Docc.count = 0 Then
		Msgbox "您尚未選取任何文件"
	End If
	
	Set doc = docc.getfirstdocument
	
	sField = Inputbox("請輸入欲修改之欄位名 : ")
	If sField="" Then	Exit Sub
%REM
	sSize = Inputbox("您所希望的增加到幾個?", "輸入","1")
	If sSize=""Then Exit Sub
%END REM
	If doc.hasitem(sField) Then 
		
		sValue = Inputbox("請輸入欲修改之欄值 : ", "輸入")
		
	End If
	
	Do While Not doc Is Nothing 
		Set doc2 = doc
'		If Not doc.hasitem(sField) Then doc.Replaceitemvalue sField,""
		If Not doc.hasitem(sField) Then
			Messagebox "無此欄位"
			End
		End If
		
		
		Call doc.replaceitemvalue(sField, sValue)
		Call doc.save(False, True)
		
		Set doc = Docc.GetNextDocument(doc2) 
		
	Loop
End Sub

