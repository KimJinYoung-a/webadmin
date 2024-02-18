<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/lookbookCls.asp"-->
<%
Dim detailidxarr, tmpSort, tmpIsusing, cnt, i, sqlStr, isusingarr, idx, mode, adminid
	detailidxarr 	= Request("detailidxarr")
	isusingarr	= Request("isusingarr")
	idx			= Request("idx")
	mode 		= Request("mode")

adminid = session("ssBctId")

if mode="sortisusingedit" then

	'선택상품 파악
	detailidxarr = split(detailidxarr,",")
	cnt = ubound(detailidxarr)
	
	isusingarr	=  split(isusingarr,",")
	
	For i = 0 to cnt
		tmpIsusing = isusingarr(i)
	
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_brand.dbo.tbl_street_LookBook_Detail SET " & VBCRLF
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'" & VBCRLF
		sqlStr = sqlStr & " ,lastupdate=getdate()" + vbcrlf
		sqlStr = sqlStr & " ,lastadminid = '"&adminid&"'" + vbcrlf		
		sqlStr = sqlStr & " WHERE detailidx =" & detailidxarr(i)
		
		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	Next

	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	location.replace('/admin/brand/lookbook/iframe_lookbook_detail.asp?idx="&idx&"');"
	response.write "</script>"

else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->