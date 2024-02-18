<% 
'========================================================================== 
'	Description: 업로드 관련 함수 모음 
'	History: 2009.02.11
'========================================================================== 
'전역변수 선언 


'## fnChkImgFile: 파일 용량 및 확장자 체크 ## 
'## input : 파일명, 업로드 최대용량 / oupput: True, False ## 

Function fnChkImgFile(ByVal sfile, ByVal smaxlen) 
	Dim  strFileSize, strFileType

	IF  sfile = "" THEN  
		fnChkImgFile = FALSE 
	ELSE	 
		strFileSize = sfile.FileSize 
		strFileType = LCase(sfile.FileType)  

		if strFileSize  > smaxlen then	'용량 체크 
			smaxlen = CLng(smaxlen)/1024
%> 
		<script language="javascript">
		<!-- 
			alert("파일크기는 <%=smaxlen%>KB이하만 가능합니다.");	 
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
			alert("JPG또는 GIF형식의 파일만 가능합니다.");		 
			history.go(-1);			 
		//--> 
		</script> 
<%			 
		response.end 
		end if 
		 
		fnChkImgFile = TRUE 
	END IF	 
End Function 
 
'## fnChkFile: 파일 용량 및 확장자 체크 ## 
'## input : 파일명, 업로드 최대용량 / oupput: True, False ## 
Function fnChkFile(ByVal sfile, ByVal smaxlen, ByVal fileType) 
	Dim  strFileSize, strFileType 
 
	IF  sfile = "" THEN  
		fnChkFile = FALSE
	ELSE	 
		strFileSize = sfile.FileSize 
		strFileType = LCase(sfile.FileType)  
	 
		if strFileSize  > smaxlen then	'용량 체크 
%> 
		<script language="javascript"> 
		<!-- 
			alert("파일크기는 <%=smaxlen%>MB이하만 가능합니다.");	 
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
			alert("<%=fileType%>형식의 파일만 가능합니다."); 
			history.go(-1);				
		//-->
		</script>
<%			
		response.end 
		end if 
		 
		fnChkFile = TRUE 
	END IF	 
End Function 

'## fnMakeFileName : 업로드파일 이름 생성  ##
'## ouput: 년월일시분초 ##
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

'## Format00: 자리수 맞추기 ## 
'## input : 원하는 자리수, 데이터 / output : '0...'+데이터 
Function Format00(ByVal n, ByVal orgData) 
    dim tmp 
	if (n-Len(CStr(orgData))) < 0 then 
		Format00 = CStr(orgData)
		Exit Function 
	end if 

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData) 
	Format00 = tmp 
End Function 

'// 접근경로 확인 // 
Sub sbCheckReferer(strWindow, strDestination) 
	If InStr(1, Request.ServerVariables("HTTP_REFERER"), strDomain, vbTextCompare) = 0 Then Call sbAlertMessage("정상적인 접근경로가 아닙니다.", strWindow,strDestination) 
End Sub 

'// 메시지 출력 및 페이지이동// 
Sub sbAlertMessage(ByVal strMessage, ByVal strWindow, ByVal strDestination) 
	'메세지 출력 
	Response.Write	"<script language='javascript'>" &_ 
							"alert('" & strMessage & "');"  

	Select Case strDestination

		'창 닫기 
		Case "close" 
			Response.Write strWindow & ".close();" 

		'이전 페이지로 
		Case "back" 
			Response.Write "history.go(-1);" 

		'해당 페이지로 이동 
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