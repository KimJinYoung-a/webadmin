<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ٿ�ε� ����  ��� 
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
'// ��庰 �б�
Select Case mode
	Case "add"			'���
		IF filesize = "" THEN 
			 Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.") 
			response.end
		END IF
		
'		/* �̺�Ʈ�ڵ� �ߺ���밡�ɿ��� üũ
'		sqlStr = " select fileSN from [db_sitemaster].[dbo].tbl_DownloadFile where evt_code ="&evt_code&" and evt_code <> 0 and evt_code is not null "  
'		rsget.Open sqlStr,dbget,1
'		IF not (rsget.EOF or rsget.BOF) THEN
'			returnvalue = rsget("fileSN")
'		END IF
'		rsget.Close 
'		IF returnvalue <> ""   THEN
'		 Call Alert_return ("�̹� ��ϵ� �̺�Ʈ�ڵ��Դϴ�. Ȯ�� �� �ٽ� �Է����ּ���") 
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

	Case "edit"			'����
	IF filesize = "" THEN 
			 Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.") 
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

	Case "delete"			'����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_DownloadFile " &_
				" Set isUsing='N' " &_
				" Where fileSn=" & fileSn
		dbget.Execute sqlStr
End Select
If Err.Number = 0 Then
	  Call Alert_move ("ó���Ǿ����ϴ�.","/admin/sitemaster/downloadFile_list.asp") 
ELSE
	 Call Alert_return ("������ ó���� ������ �߻��Ͽ����ϴ�.") 
END IF	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->