'(�����ĵ�����������ϵͳ)|SaveToLocalSystem: 

Option Public

Type BROWSEINFO 
	hOwner As Long 
	pidlRoot As Long 
	pszDisplayName As String 
	lpszTitle As String 
	ulFlags As Long 
	lpfn As Long 
	lParam As Long 
	iImage As Long 
End Type 

Const BIF_RETURNONLYFSDIRS = &H1 
Const BIF_DONTGOBELOWDOMAIN = &H2 
Const BIF_STATUSTEXT = &H4 
Const BIF_RETURNFSANCESTORS = &H8 
Const BIF_BROWSEFORCOMPUTER = &H1000 
Const BIF_BROWSEFORPRINTER = &H2000 
Const MAX_PATH = 260 

Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (Byval pidl As Long, Byval pszPath As String) As Long 
Declare Function SHBrowseForFolder Lib "shell32"(lpBrowseInfo As BROWSEINFO) As Long 

Declare Sub CoTaskMemFree Lib "ole32" (Byval pv As Long) 
Declare Function GetDesktopWindow Lib "user32" () As Long 

'�Զ���ȫ�ֱ���
Const rtfName="fldDocBody" 'rtf�ֶ�����
Const categoryName="fldCateNamePath" '�����ֶ�
Const seperator="\"  '�㼶Ŀ¼�������ֶηָ����Windowsƽ̨ʹ��"\",Unixʹ��"/"
Sub Initialize() 
	'/***
	' @Date 	2013-10-24
	' @Author	�����	
	'  		
	' �����������������ĳ�������ͼ��ѡ���ĵ������ĵ�ĳ��RTF�ֶ��еĸ������ڱ����ļ����У�
	' ���������ͼ��һ�еķ��ഴ��Ŀ¼�㼶�����ļ�������Ӧ��Ŀ¼�㼶��
	' ����:��ͼ��һ�з���Ϊ"AA/11",�����кܶ��ĵ�����ô��Ӧ�����ͻ����/AA/11Ŀ¼��
	' ***********************/
	On Error Goto ErrHandler
	
	Dim session As New NotesSession 
	Dim db As NotesDatabase 
	Dim collection As NotesDocumentCollection 
	Dim doc As NotesDocument   
	Dim rtitem As Variant 
	Dim NotesItem As NotesItem 
	
	Dim bi As BROWSEINFO 
	Dim pidl As Long 
	Dim path As String 
	Dim pos As Integer 
	
	bi.hOwner = GetDesktopWindow() 
	bi.pidlRoot = 0& 
	bi.lpszTitle = "Select directory to save the attachments"
	bi.ulFlags = BIF_RETURNONLYFSDIRS 
	pidl = SHBrowseForFolder(bi) 
	
	If pidl=0 Then
		Exit Sub '�����ȡ����ť
	End If
	
	path = Space$(MAX_PATH) 
	
	If SHGetPathFromIDList(Byval pidl, Byval path) Then 
		pos = Instr(path, Chr$(0)) 
	End If 
	
	Call CoTaskMemFree(pidl) 
	
	Set db = session.CurrentDatabase 
	Set collection = db.UnprocessedDocuments 
	Set doc = collection.GetFirstDocument() 
	
	Dim orgDir As String
	orgDir= Left(path, pos - 1)
	Dim desDir As String
	Dim desDirVar As Variant
	
	Print "(�����ĵ�����������ϵͳ)|SaveToLocalSystem:������ָ��Ŀ¼�б��渽��,�����ĵȴ�..."
	
	While Not(doc Is Nothing)   
     ' �˴��ٶ�������Ƕ���� Body ���У���ȻҲ����ѭ���ĵ����е���Ȼ����ڸ��ı�����д�����ȡ���� 
		
		Chdir(orgdir) '�ı䵽��ʼĿ¼
		desDir=doc.GetItemValue(categoryName)(0)
		desDirVar=Split(desDir,seperator)
		
		Forall x In desDirVar
			If Dir(x,16) = "" Then  '�Ƿ����Ŀ¼
				Mkdir(x)           '�������򴴽�Ŀ¼  
			End If
			Chdir(x)				'����Ŀ¼�󣬶�λ����Ŀ¼						
		End Forall
		
		
		Set rtitem = doc.GetFirstItem( rtfName ) 
		If ( rtitem.Type = RICHTEXT ) Then 
			Forall o In rtitem.EmbeddedObjects                     
				If ( o.Type = EMBED_ATTACHMENT ) Then     
					'Call o.ExtractFile( Left(path, pos - 1) & "\" & o.Name  ) 
					Call o.ExtractFile(o.Name) 
				End If         
			End Forall 
		End If 
		
		Set doc = collection.GetNextDocument(doc) 
	Wend 
	
agEnd:	
	Chdir(orgdir)  'ָ��ָ���Ǹ�Ŀ¼,�Ǹ�Ŀ¼���޷���ɾ��
	Print db.Title + ":" + "����֮(�����ĵ�����������ϵͳ)|SaveToLocalSystem ����ִ�С�"
	Exit Sub
	
ErrHandler	:
	Print "(�����ĵ�����������ϵͳ)|SaveToLocalSystem:" & Error & " " & Erl & "**" & Err
	Goto agEnd
End Sub

