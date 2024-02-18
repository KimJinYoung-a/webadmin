<% 
'========================================================================== 
'	Description: ���ε� ���� �Լ� ���� 
'	History: 2009.02.11
'========================================================================== 
'�������� ���� 


'## fnChkImgFile: ���� �뷮 �� Ȯ���� üũ ## 
'## input : ���ϸ�, ���ε� �ִ�뷮 / oupput: True, False ## 

Function fnChkImgFile(ByVal sfile, ByVal smaxlen) 
	Dim  strFileSize, strFileType

	IF  sfile = "" THEN  
		fnChkImgFile = FALSE 
	ELSE	 
		strFileSize = sfile.FileSize 
		strFileType = LCase(sfile.FileType)  

		if strFileSize  > smaxlen then	'�뷮 üũ 
			smaxlen = CLng(smaxlen)/1024
%> 
		<script language="javascript">
		<!-- 
			alert("����ũ��� <%=smaxlen%>KB���ϸ� �����մϴ�.");	 
			history.go(-1);				 
		//--> 
		</script>	 
<%			 
		response.end 
		end if 
		 
		if not (  strFileType = "gif" or strFileType = "jpeg" or strFileType = "jpg" ) then 
%> 
		<script language="javascript">
		<!-- 
			alert("JPG�Ǵ� GIF������ ���ϸ� �����մϴ�.");		 
			history.go(-1);			 
		//--> 
		</script> 
<%			 
		response.end 
		end if 
		 
		fnChkImgFile = TRUE 
	END IF	 
End Function 
 
'## fnChkFile: ���� �뷮 �� Ȯ���� üũ ## 
'## input : ���ϸ�, ���ε� �ִ�뷮 / oupput: True, False ## 
Function fnChkFile(ByVal sfile, ByVal smaxlen, ByVal fileType) 
	Dim  strFileSize, strFileType 
 
	IF  sfile = "" THEN  
		fnChkFile = FALSE
	ELSE	 
		strFileSize = sfile.FileSize 
		strFileType = LCase(sfile.FileType)  
	 
		if strFileSize  > smaxlen then	'�뷮 üũ 
%> 
		<script language="javascript"> 
		<!-- 
			alert("����ũ��� <%=smaxlen%>MB���ϸ� �����մϴ�.");	 
			history.go(-1);		 
		//--> 
		</script>	 
<%			 
		response.end 
		end if 
		 
		if not (  strFileType = fileType ) then 
%> 
		<script language="javascript"> 
		<!-- 
			alert("<%=fileType%>������ ���ϸ� �����մϴ�."); 
			history.go(-1);				
		//-->
		</script>
<%			
		response.end 
		end if 
		 
		fnChkFile = TRUE 
	END IF	 
End Function 

'## fnMakeFileName : ���ε����� �̸� ����  ##
'## ouput: ����Ͻú��� ##
Function fnMakeFileName(ByVal strFile)	
	fnMakeFileName = fnMakeDateFrm&"."& strFile.FileType
End Function	

Function fnMakeDateFrm 
	Dim sNow, sY, sM, sD, sH, sMi, sS 
	sNow = now() 
	sY= Year(sNow) 
	sM = Format00(2,Month(sNow)) 
	sD = Format00(2,Day(sNow)) 
	sH = Format00(2,Hour(sNow))
	sMi = Format00(2,Minute(sNow)) 
	sS = Format00(2,Second(sNow)) 
	fnMakeDateFrm = sY&sM&sD&sH&sMi&sS 
End Function 

Function fnMakeDateFolderName 
	Dim sNow, sY, sM
	sNow = now() 
	sY= Year(sNow) 
	sM = Format00(2,Month(sNow))  
	fnMakeDateFolderName = sY&sM
End Function 

'## Format00: �ڸ��� ���߱� ## 
'## input : ���ϴ� �ڸ���, ������ / output : '0...'+������ 
Function Format00(ByVal n, ByVal orgData) 
    dim tmp 
	if (n-Len(CStr(orgData))) < 0 then 
		Format00 = CStr(orgData)
		Exit Function 
	end if 

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData) 
	Format00 = tmp 
End Function 

'// ���ٰ�� Ȯ�� // 
Sub sbCheckReferer(strWindow, strDestination) 
	If InStr(1, Request.ServerVariables("HTTP_REFERER"), strDomain, vbTextCompare) = 0 Then Call sbAlertMessage("�������� ���ٰ�ΰ� �ƴմϴ�.", strWindow,strDestination) 
End Sub 

'// �޽��� ��� �� �������̵�// 
Sub sbAlertMessage(ByVal strMessage, ByVal strWindow, ByVal strDestination) 
	'�޼��� ��� 
	Response.Write	"<script language='javascript'>" &_ 
							"alert('" & strMessage & "');"  

	Select Case strDestination

		'â �ݱ� 
		Case "close" 
			Response.Write strWindow & ".close();" 

		'���� �������� 
		Case "back" 
			Response.Write "history.go(-1);" 

		'�ش� �������� �̵� 
		Case Else 
			Response.Write strWindow & ".location.href='" & strDestination & "';" 

	End Select 

	IF strWindow = "opener" THEN Response.Write  "self.close();" 
	Response.Write "</script>" 
	Response.End 
	 
End Sub 

function html2db(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function
	v = Replace(v, "'", "''")
	html2db = v
end Function
%> 