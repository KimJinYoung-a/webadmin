<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/shopcls.asp"-->
<%
Dim detailidxarr, sortnoarr, tmpSort, tmpIsusing, cnt, i, sqlStr, isusingarr, idx, mode, adminid
	sortnoarr 	= Request("sortnoarr")
	detailidxarr 	= Request("detailidxarr")
	isusingarr	= Request("isusingarr")
	idx			= Request("idx")
	mode 		= Request("mode")

adminid = session("ssBctId")

if mode="sortisusingedit" then
	If sortnoarr="" or idx="" THEN
		Response.Write "<script language='javascript'>alert('구분자가 없습니다'); history.back(-1);</script>"
		dbget.close()	:	response.End
	end if
	
	'선택상품 파악
	detailidxarr = split(detailidxarr,",")
	cnt = ubound(detailidxarr)
	
	sortnoarr 	=  split(sortnoarr,",")
	isusingarr	=  split(isusingarr,",")
	
	For i = 0 to cnt
		tmpSort = sortnoarr(i)	
		tmpIsusing = isusingarr(i)
	
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_brand.dbo.tbl_street_shop_collection_item SET " & VBCRLF
		sqlStr = sqlStr & " sortNo = '"&tmpSort&"'" & VBCRLF
		sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'" & VBCRLF
		sqlStr = sqlStr & " ,lastupdate=getdate()" + vbcrlf
		sqlStr = sqlStr & " ,lastadminid = '"&adminid&"'" + vbcrlf		
		sqlStr = sqlStr & " WHERE detailidx =" & detailidxarr(i)
		
		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	Next

	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	parent.location.reload();"
	response.write "</script>"

else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->