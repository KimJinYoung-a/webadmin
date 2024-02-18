<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : docatetag.asp
' Discription : appcatetag 처리 페이지
' History : 2014-09-02 이종화 생성
'###############################################

'// 변수 선언 및 파라메터 접수
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
		'신규 등록
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

		Response.write "<script>alert('신규 등록 완료.');</script>"
		Response.write "<script>location.href='http://webadmin.10x10.co.kr/admin/mobile/catetag/?menupos="&menupos&"';</script>"
	Else
		'내용 수정
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

		Response.write "<script>alert('수정 완료');</script>"
		Response.write "<script>location.href='http://webadmin.10x10.co.kr/admin/mobile/catetag/?menupos="&menupos&"';</script>"
	End If

	'// 목록으로 복귀
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
