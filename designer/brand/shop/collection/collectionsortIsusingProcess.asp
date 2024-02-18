<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/shopcls.asp"-->
<%
Dim detailidxarr, sortnoarr, tmpSort, tmpIsusing, cnt, i, sqlStr, isusingarr, idx, mode, adminid
dim tmpstate
	sortnoarr 	= Request("sortnoarr")
	detailidxarr 	= Request("detailidxarr")
	isusingarr	= Request("isusingarr")
	idx			= requestCheckVar(Request("idx"),10)
	mode 		= requestCheckVar(Request("mode"),30)

adminid = session("ssBctId")

if mode="sortisusingedit" then
	If sortnoarr=""  or idx="" THEN
		Response.Write "<script language='javascript'>alert('구분자가 없습니다'); history.back(-1);</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "SELECT top 1 state"
	sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_shop_collection as M"
	sqlStr = sqlStr & " WHERE m.idx="&idx&""
	
	'response.write sqlStr & "<BR>"
	rsget.Open sqlStr, dbget, 1
    If Not rsget.Eof then
    	tmpstate = rsget("state")
	End If
    rsget.Close

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

	if tmpstate="7" then
		sqlStr = "UPDATE db_brand.dbo.tbl_street_shop_collection SET" + VBCRLF
		sqlStr = sqlStr & " state = 2" + VBCRLF
		sqlStr = sqlStr & " where idx ='" & Cstr(idx) & "'"

		'response.write sqlStr & "<BR>"	
		dbget.execute sqlStr
	end if
	
	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	parent.location.reload();"	
	response.write "</script>"

else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->