'�ֶ�\�ֶ��滻�ֶ���������(AllReplace): 

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
	
	' �ڹ��ܱ����Ƿ��Ј��Й���
	Dim level As Integer
	level = db.CurrentAccessLevel
	If level < ACLLEVEL_DESIGNER Then
		Messagebox "���o�����д˳�ʽ!",0,"����"
		Exit Sub
	End If
	
	If Docc.count = 0 Then
		Msgbox "����δ�xȡ�κ��ļ�"
	End If
	
	Set doc = docc.getfirstdocument
	
	sField = Inputbox("Ոݔ�����޸�֮��λ�� : ")
	If sField="" Then	Exit Sub
%REM
	sSize = Inputbox("����ϣ�������ӵ��ׂ�?", "ݔ��","1")
	If sSize=""Then Exit Sub
%END REM
	If doc.hasitem(sField) Then 
		
		sValue = Inputbox("Ոݔ�����޸�֮��ֵ : ", "ݔ��")
		
	End If
	
	Do While Not doc Is Nothing 
		Set doc2 = doc
'		If Not doc.hasitem(sField) Then doc.Replaceitemvalue sField,""
		If Not doc.hasitem(sField) Then
			Messagebox "�o�˙�λ"
			End
		End If
		
		
		Call doc.replaceitemvalue(sField, sValue)
		Call doc.save(False, True)
		
		Set doc = Docc.GetNextDocument(doc2) 
		
	Loop
End Sub

