'(保存文档附件到本地系统)|SaveToLocalSystem: 

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

'自定义全局变量
Const rtfName="fldDocBody" 'rtf字段名字
Const categoryName="fldCateNamePath" '分类字段
Const seperator="\"  '层级目录产生的字段分割符，Windows平台使用"\",Unix使用"/"
Sub Initialize() 
	'/***
	' @Date 	2013-10-24
	' @Author	Alex Yean	
	'  		
	' 这是批量附件导出的程序：在视图中选择文档，把文档某个RTF字段中的附件放于本地文件夹中；
	' 并会根据视图第一列的分类创建目录层级，把文件放于相应的目录层级中
	' 例如:视图第一列分类为"AA/11",其下有很多文档，那么对应附件就会放在/AA/11目录中
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
		Exit Sub '点击了取消按钮
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
	
	Print "(保存文档附件到本地系统)|SaveToLocalSystem:正在向指定目录中保存附件,请耐心等待..."
	
	While Not(doc Is Nothing)   
     ' 此处假定附件是嵌入在 Body 域当中，当然也可以循环文档所有的域，然后对于富文本域进行处理，提取附件 
		
		Chdir(orgdir) '改变到初始目录
		desDir=doc.GetItemValue(categoryName)(0)
		desDirVar=Split(desDir,seperator)
		
		Forall x In desDirVar
			If Dir(x,16) = "" Then  '是否存在目录
				Mkdir(x)           '不存在则创建目录  
			End If
			Chdir(x)				'存在目录后，定位到该目录						
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
	Chdir(orgdir)  '指针指在那个目录,那个目录就无法被删除
	Print db.Title + ":" + "代理之(保存文档附件到本地系统)|SaveToLocalSystem 结束执行。"
	Exit Sub
	
ErrHandler	:
	Print "(保存文档附件到本地系统)|SaveToLocalSystem:" & Error & " " & Erl & "**" & Err
	Goto agEnd
End Sub

