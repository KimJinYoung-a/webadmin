<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  �ΰŽ� ��ī���� PC���� �۰�&���� ��ũ �Է�,���� ó�� ������
' History : 2016-10-24 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, startdate , titletext, contentstext, lectureid  , mode, isusing

idx			= RequestCheckVar(request("idx"),10)
mode		= RequestCheckVar(request("mode"),10)
isusing		= requestCheckvar(request("isusing"),2)
titletext	= RequestCheckVar(request("titletext"),100)
startdate	= RequestCheckVar(request("startdate"),10)
lectureid	= RequestCheckVar(request("lectureid"),32)
contentstext	= RequestCheckVar(request("contentstext"),100)
  	if titletext <> "" then
		if checkNotValidHTML(titletext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if contentstext <> "" then
		if checkNotValidHTML(contentstext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
if idx = "" then
	idx = 0
end If

if idx = 0 then
	''�ű� ���
	mode = "add"
else
	''����
	mode = "edit"
end if

dim sqlStr

if (mode = "add") then
''�ű� ���
    sqlStr = " insert into [db_academy].[dbo].tbl_academy_PCmain_lectureLink" + VbCrlf
    sqlStr = sqlStr + " (lectureid, titletext, contentstext, startdate, isusing)" + VbCrlf
    sqlStr = sqlStr + " values(" + VbCrlf
    sqlStr = sqlStr + " '" + lectureid + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + html2db(titletext) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + html2db(contentstext) + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + startdate + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " )"
'	response.write sqlStr
'	response.end
    dbACADEMYget.Execute sqlStr

elseif mode = "edit" Then
''����
   sqlStr = " update  [db_academy].[dbo].tbl_academy_PCmain_lectureLink " + VbCrlf
   sqlStr = sqlStr + " set " + VbCrlf
   sqlStr = sqlStr + " lectureid='" + lectureid + "'" + VbCrlf
   sqlStr = sqlStr + " ,titletext='" + html2db(titletext) + "'" + VbCrlf
   sqlStr = sqlStr + " ,contentstext='" + html2db(contentstext) + "'" + VbCrlf
   sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
   sqlStr = sqlStr + " ,startdate='" + startdate + "'" + VbCrlf
   sqlStr = sqlStr + " where idx=" + CStr(idx)
   dbACADEMYget.Execute sqlStr
end if
%>
<script language = "javascript">
	alert("����Ǿ����ϴ�.");
	opener.location.reload();
	self.close();
</script>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->