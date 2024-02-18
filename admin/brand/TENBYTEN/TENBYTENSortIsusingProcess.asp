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

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/TENBYTENCls.asp"-->
<%
Dim itemidarr, sortarr, tmpSort, tmpIsusing, cnt, i, sqlStr, isusingarr, makerid
	sortarr 	= Request("sortarr")
	itemidarr 	= Request("itemidarr")
	isusingarr	= Request("isusingarr")
	menupos	= request("menupos")

If sortarr="" THEN
	Response.Write "<script language='javascript'>history.back(-1);</script>"
	dbget.close()	:	response.End
end if

'선택상품 파악
itemidarr = split(itemidarr,",")
cnt = ubound(itemidarr)

sortarr 	=  split(sortarr,",")
isusingarr	=  split(isusingarr,",")

For i = 0 to cnt
	tmpSort = sortarr(i)	
	tmpIsusing = isusingarr(i)

	sqlStr = "UPDATE db_brand.dbo.tbl_street_TENBYTEN SET " & VBCRLF
	sqlStr = sqlStr & " sortNo = '"&tmpSort&"'" & VBCRLF
	sqlStr = sqlStr & " ,isusing = '"&tmpIsusing&"'" & VBCRLF
	sqlStr = sqlStr & " WHERE idx =" & itemidarr(i)
	
	'response.write sqlStr & "<br>"
	dbget.execute sqlStr
Next

%>

<script language='javascript'>
	alert('저장되었습니다');
	location.replace('/admin/brand/TENBYTEN/index.asp?makerid=<%= makerid %>&menupos=<%= menupos %>');
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->