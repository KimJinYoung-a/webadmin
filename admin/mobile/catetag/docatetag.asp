<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : docatetag.asp
' Discription : appcatetag ó�� ������
' History : 2014-09-02 ����ȭ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim idx , strMsg , tmpeCode , sqlStr
Dim kword1 , appdiv , kwordurl1 , kwordurl2 , isusing , catecode
Dim mode , menupos , appcate

	mode		= Request("mode")
	idx			= Request("idx")
	appdiv		= Request("appdiv")
	isusing		= Request("isusing")
	catecode	= Request("disp")
	menupos		= Request("menupos")
	kword1		= Request("kword1")
	kwordurl1	= Request("kwordurl1")
	kwordurl2	= Request("kwordurl2")
	appcate		= Request("appcate")

	If mode = "add" then
		'�ű� ���
		sqlStr = "Insert Into db_sitemaster.dbo.tbl_mobile_catetag " &_
					" (appdiv , kword1 , kwordurl1 , kwordurl2 , catecode , appcate) values " &_
					" ('" & appdiv &"'" &_
					" ,'" & kword1 &"'" &_
					" ,'" & kwordurl1 &"'" &_
					" ,'" & kwordurl2 &"'" &_
					" ,'" & catecode &"'" &_
					" ,'" & appcate &"'" &_
					")"
		dbget.Execute(sqlStr)

		Response.write "<script>alert('�ű� ��� �Ϸ�.');</script>"
		Response.write "<script>location.href='http://webadmin.10x10.co.kr/admin/mobile/catetag/?menupos="&menupos&"';</script>"
	Else
		'���� ����
		sqlStr = "Update db_sitemaster.dbo.tbl_mobile_catetag " &_
				" Set appdiv='" & appdiv & "'" &_
				" 	,kword1='" & kword1 & "'" &_
				" 	,kwordurl1='" & kwordurl1 & "'" &_
				" 	,kwordurl2='" & kwordurl2 & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,catecode='" & catecode & "'" &_
				" 	,appcate='" & appcate & "'" &_
				" 	,lastupdate=getdate()" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

		Response.write "<script>alert('���� �Ϸ�');</script>"
		Response.write "<script>location.href='http://webadmin.10x10.co.kr/admin/mobile/catetag/?menupos="&menupos&"';</script>"
	End If

	'// ������� ����
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
