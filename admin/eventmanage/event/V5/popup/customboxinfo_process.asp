<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : customboxinfo_process.asp
' Discription : I��(������) �̺�Ʈ ����� ���� �ڽ� ��� ���μ���
' History : 2019.02.13 ������
'###############################################
session.codePage = 65001		'�����ڵ� UTF-8 ���� ����
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, sqlStr, title, menuidx, customcontents, idx
dim refer, BGImage, BGColorLeft, BGColorRight, contentsAlign, Margin
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
menuidx = requestCheckVar(Request.Form("menuidx"),10)
idx = requestCheckVar(Request.Form("idx"),10)
title = requestCheckVar(Request.Form("title"),64)
customcontents = html2db(Request.Form("customcontents"))
BGImage	= requestCheckVar(Request.form("BGImage"),128)
BGColorLeft	= requestCheckVar(Request.form("BGColorLeft"),8)
BGColorRight	= requestCheckVar(Request.form("BGColorRight"),8)
contentsAlign	= requestCheckVar(Request.form("contentsAlign"),1)
Margin	= requestCheckVar(Request.form("Margin"),10)

if BGColorLeft="" then BGColorLeft="#FFFFFF"

if title <> "" then
	if checkNotValidHTML(title) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	session.codePage = 949		'�����ڵ� EUC-KR ����
	response.End
	end if
end If

if customcontents <> "" then
	if checkNotValidHTML(customcontents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	session.codePage = 949		'�����ڵ� EUC-KR ����
	response.End
	end if
end If

if eCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ������ �Դϴ�. �ٽ� �õ��� �ּ���.');history.back();"
	response.write "</script>"
	session.codePage = 949		'�����ڵ� EUC-KR ����
	response.End
end if

if BGImage <> "" then
	if checkNotValidHTML(BGImage) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
	if not isNumeric(Margin) then Margin=0
    '// ��Ƽ������ ������ ���� �Է�
    sqlStr = " Update db_event.dbo.tbl_event_multi_contents_master" & vbCrLf
    sqlStr = sqlStr & " Set BGImage='" & BGImage & "'" & vbCrLf
    sqlStr = sqlStr & " ,BGColorLeft='" & BGColorLeft & "'" & vbCrLf
	sqlStr = sqlStr & " ,BGColorRight='" & BGColorRight & "'" & vbCrLf
    sqlStr = sqlStr & " ,contentsAlign='" & contentsAlign & "'" & vbCrLf
    sqlStr = sqlStr & " ,Margin='" & Margin & "'" & vbCrLf
    sqlStr = sqlStr & " Where idx='" & menuidx & "'"
    dbget.Execute sqlStr

    '--3.theme ����
    sqlStr = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
    sqlStr = sqlStr + " SET contentsAlign='" & contentsAlign & "'" & vbCrlf
    sqlStr = sqlStr + " where evt_code=" & eCode
    dbget.execute sqlStr

select case eMode
case "CI"
	dbget.beginTrans
    '===========================================================
    '-- �߰�
		sqlStr = "Insert Into db_event.dbo.tbl_event_multi_contents " & vbcrlf
		sqlStr = sqlStr + " (menuidx, title, BrandContents) values "  & vbcrlf
		sqlStr = sqlStr + " ('" & menuidx  & "'"  & vbcrlf
		sqlStr = sqlStr + " ,'" & title &"'"  & vbcrlf
		sqlStr = sqlStr + " ,'" & customcontents &"')"
		dbget.Execute(sqlStr)

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
			session.codePage = 949		'�����ڵ� EUC-KR ����
            response.End
        end if
    '===========================================================
	dbget.CommitTrans

	response.write "<script type='text/javascript'>"
	response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
	response.write "    self.close();"
	response.write "</script>"
	session.codePage = 949		'�����ڵ� EUC-KR ����
	dbget.close()	:	response.End
case "CU"
	dbget.beginTrans
    '===========================================================
    '-- ����
        sqlStr = "UPDATE [db_event].[dbo].[tbl_event_multi_contents]" & vbCrlf
        sqlStr = sqlStr + " SET title='" & title & "'" & vbCrlf
        sqlStr = sqlStr + " , BrandContents='" & customcontents & "'" & vbCrlf
        sqlStr = sqlStr + " where idx=" & idx
        dbget.execute sqlStr

        if Err.Number <> 0 then
            dbget.RollBackTrans
            Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[2]", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans

	response.write "<script type='text/javascript'>"
	response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
	response.write "    self.close();"
	response.write "</script>"
	session.codePage = 949		'�����ڵ� EUC-KR ����
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
	session.codePage = 949		'�����ڵ� EUC-KR ����
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->