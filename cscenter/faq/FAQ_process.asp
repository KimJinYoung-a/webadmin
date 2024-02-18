<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]��������>>[FAQ]���� 
' Hieditor : 2009.03.02 �̿��� ����
'			 2021.07.30 �ѿ�� ����(��뿩�� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/faq_cls.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim faqid, userid, regusername
dim linkname, linkurl, disporder
dim title, contents, commCd
dim SQL
dim page, searchDiv, searchKey, searchString, param, retURL, isusing
	isusing = requestcheckvar(request("isusing"),1)

'// ���� ���� �� ó��
menupos		= Request("menupos")
faqid		= Request("faqid")
mode		= Request("mode")
commCd		= Request("commCd")
title		= html2db(Request("title"))
contents	= html2db(Request("contents"))

linkname    = html2db(Request("linkname"))
linkurl     = html2db(Request("linkurl"))
disporder   = Request("disporder")

page		= Request("page")
searchDiv	= Request("searchDiv")
searchKey	= Request("searchKey")
searchString = Request("searchString")

param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����


Dim mode		: mode		= req("mode","INS")
Dim i, PKID

Dim obj	: Set obj = new Cfaq

obj.FRectfaqid = ""
obj.GetFAQRead

obj.FfaqList(0).FfaqID				= req("faqID","")
obj.FfaqList(0).FcommCd				= req("commCd","")
obj.FfaqList(0).FdispOrder			= req("dispOrder",999)
obj.FfaqList(0).Ftitle				= ReplaceBracket(req("title",""))
obj.FfaqList(0).Fcontents			= ReplaceBracket(req("contents",""))
obj.FfaqList(0).FlinkName			= ReplaceBracket(req("linkName",""))
obj.FfaqList(0).FlinkUrl				= ReplaceBracket(req("linkUrl",""))
obj.FfaqList(0).fisusing				= req("isusing","")

Dim ErrMsg
If mode = "DEL" Or mode = "USE" Then	' ����, ���
	PKID = Split(Replace(req("faqID","")," ",""),",")
	For i = 0 To UBound(PKID)
		obj.FfaqList(0).FfaqID		= PKID(i)
		ErrMsg = obj.ProcData(mode)
	Next 
Else					' ���,����
	ErrMsg = obj.ProcData(mode)
End If 

Set obj = Nothing 


If ErrMsg <> "" Then 
	response.write	"<script language='javascript'>" &_
					"	alert('" & ErrMsg & "');" &_
					"	history.back();" &_
					"</script>"
Else 
	If mode = "UPD" Then 
		retURL = "faq_view.asp?menupos=" & menupos & "&faqid=" & faqid & param
	Else 
		retURL = "faq_list.asp?menupos=" & menupos & param
	End If 
	response.write	"<script language='javascript'>" &_
					"	alert('" & getModeName(mode) & "�Ǿ����ϴ�.');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
End If 

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->