<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################################
' PageName : pop_login_mileage.asp
' Discription : I��(������) �̺�Ʈ ������ �α��� ���ϸ��� ���� ���
' History : 2021.11.26 ������
'###############################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->

<%

dim evt_code, mileagePoint, jukyo, existsCheck, sqlStr

evt_code = requestCheckVar(Request.Form("evt_code"),10)
mileagePoint = requestCheckVar(Request.Form("mileagePoint"),10)
jukyo = requestCheckVar(Request.Form("jukyo"),128)

if jukyo <> "" then
	if checkNotValidHTML(jukyo) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if evt_code="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.');history.back();"
	response.write "</script>"
	response.End
end if

	sqlStr = " SELECT top 1 evt_code" + vbcrlf
	sqlStr = sqlStr + " FROM [db_event].[dbo].[tbl_event_login_mileage]" + vbcrlf
	sqlStr = sqlStr + " WHERE evt_code=" & evt_code
	rsget.Open sqlStr, dbget
	IF not rsget.EOF THEN
		existsCheck = rsget("evt_code")
	END IF	
	rsget.close	

	if existsCheck > 0 then
		sqlStr = " Update [db_event].[dbo].[tbl_event_login_mileage]" & vbCrLf
		sqlStr = sqlStr & " Set mileage='" & mileagePoint & "'" & vbCrLf
		sqlStr = sqlStr & " ,jukyo='" & jukyo & "'" & vbCrLf
		sqlStr = sqlStr & " ,lastupdate=getdate()" & vbCrLf
		sqlStr = sqlStr & " Where evt_code='" & evt_code & "'"
		dbget.Execute sqlStr
	else
		sqlStr =" insert into [db_event].[dbo].[tbl_event_login_mileage]" & VbCrlf
		sqlStr = sqlStr & " (evt_code, mileage, jukyo, reguserid)" & VbCrlf
		sqlStr = sqlStr & " values(" & CStr(evt_code) & "," & mileagePoint & ",'" & jukyo & "','" & session("ssBctId") & "')" & VbCrlf
		dbget.execute sqlStr
	end if

	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(evt_code) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->