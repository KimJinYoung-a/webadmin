<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/lookbookCls.asp"-->
<%
Dim detailidxarr, tmpSort, tmpIsusing, cnt, i, sqlStr, isusingarr, idx, mode, adminid
dim tmpstate
	detailidxarr 	= Request("detailidxarr")
	isusingarr	= Request("isusingarr")
	idx			= requestCheckVar(Request("idx"),10)
	mode 		= requestCheckVar(Request("mode"),30)

adminid = session("ssBctId")

if mode="sortisusingedit" then
	sqlStr = "SELECT top 1 state"
	sqlStr = sqlStr & " FROM db_brand.dbo.tbl_street_LookBook_Master as M"
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

	if tmpstate="7" then
		sqlStr = "UPDATE db_brand.dbo.tbl_street_LookBook_Master SET" + VBCRLF
		sqlStr = sqlStr & " state = 2" + VBCRLF
		sqlStr = sqlStr & " where idx ='" & Cstr(idx) & "'"

		'response.write sqlStr & "<BR>"	
		dbget.execute sqlStr
	end if
	
	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	parent.location.replace('/designer/brand/lookbook/lookbookModify.asp?idx="&idx&"');"
	response.write "</script>"

else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->