<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 다운로드 파일  등록 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
Dim mode
Dim fileTitle, fileDownNm, userid,fileSn,filename,filesize,evt_code 
Dim sqlStr
mode = requestCheckvar(Request("mode"),10)
fileSn = requestCheckvar(Request("fileSn"),10)
fileTitle = HTML2DB(requestCheckvar(Request("fileTitle"),32))
fileDownNm = HTML2DB(requestCheckvar(Request("fileDownNm"),32))
filename	=requestCheckvar(Request("fName"),128)
filesize	=requestCheckvar(Request("fSize"),10)
evt_code=requestCheckvar(Request("iEC"),10)
userid = session("ssBctId")
IF evt_code = "" THEN evt_code =0
	dim returnvalue
'// 모드별 분기
Select Case mode
	Case "add"			'등록
		IF filesize = "" THEN 
			 Call Alert_return ("데이터 처리에 문제가 발생하였습니다.") 
			response.end
		END IF
		
'		/* 이벤트코드 중복사용가능여부 체크
'		sqlStr = " select fileSN from [db_sitemaster].[dbo].tbl_DownloadFile where evt_code ="&evt_code&" and evt_code <> 0 and evt_code is not null "  
'		rsget.Open sqlStr,dbget,1
'		IF not (rsget.EOF or rsget.BOF) THEN
'			returnvalue = rsget("fileSN")
'		END IF
'		rsget.Close 
'		IF returnvalue <> ""   THEN
'		 Call Alert_return ("이미 등록된 이벤트코드입니다. 확인 후 다시 입력해주세요") 
'		response.end
'		END IF 
'	*/ 
		
		sqlStr = "Insert into [db_sitemaster].[dbo].tbl_DownloadFile (fileTitle, fileDownNm,filename,filesize, userid, evt_code) values " &_
				" ('" & fileTitle & "'" &_
				" ,'" & fileDownNm & "'" &_
				" ,'" & filename & "'" &_
				" ," & filesize&_
				" ,'" & userid & "'" &_
				" , " & evt_code & " )"  
		dbget.Execute sqlStr 

	Case "edit"			'수정
	IF filesize = "" THEN 
			 Call Alert_return ("데이터 처리에 문제가 발생하였습니다.") 
			response.end
		END IF
  
		sqlStr = "Update [db_sitemaster].[dbo].tbl_DownloadFile " &_
				" Set fileTitle='" & fileTitle & "' " &_
				" 	, fileDownNm='" & fileDownNm & "' " &_
				" 	, filename='" & filename & "' " &_
				" 	, filesize= " & filesize  &_
				" 	, evt_code= " & evt_code  &_
				" 	,userid ='" & userid & "'" &_
				" Where fileSn=" & fileSn
		dbget.Execute sqlStr

	Case "delete"			'삭제
		sqlStr = "Update [db_sitemaster].[dbo].tbl_DownloadFile " &_
				" Set isUsing='N' " &_
				" Where fileSn=" & fileSn
		dbget.Execute sqlStr
End Select
If Err.Number = 0 Then
	  Call Alert_move ("처리되었습니다.","/admin/sitemaster/downloadFile_list.asp") 
ELSE
	 Call Alert_return ("데이터 처리에 문제가 발생하였습니다.") 
END IF	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->