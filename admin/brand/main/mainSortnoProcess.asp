<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.10.15 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim fidx, sortnoarr, cnt, i, tmpSort, sqlStr, mode
sortnoarr 	= Request("sortnoarr")
fidx	 	= Request("fidx")
menupos		= request("menupos")
mode		= Request("mode")

If sortnoarr="" THEN
	Response.Write "<script language='javascript'>alert('순서가 지정되지 않았습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End
End If

'선택상품 파악
fidx = split(fidx,",")
cnt = ubound(fidx) - 1

'// 정렬순서 저장
If sortnoarr<>"" THEN
	sortnoarr =  split(sortnoarr,",")

	For i = 0 to cnt
		IF sortnoarr(i) = "" THEN
			 tmpSort = "0"				
		ELSE	
			 tmpSort = sortnoarr(i)	
		END IF

		If mode = "interview" Then
			sqlStr = "UPDATE db_brand.dbo.tbl_street_interview_main SET" + vbcrlf
			sqlStr = sqlStr & " mainSortNo = "&tmpSort&"" + vbcrlf
			sqlStr = sqlStr & " WHERE mainidx =" + fidx(i)
			dbget.execute sqlStr
		ElseIf mode = "lookbook" Then
			sqlStr = "UPDATE db_brand.dbo.tbl_street_LookBook_Master SET" + vbcrlf
			sqlStr = sqlStr & " mainpageSortNo = "&tmpSort&"" + vbcrlf
			sqlStr = sqlStr & " WHERE idx =" + fidx(i)
			dbget.execute sqlStr
		Else
			sqlStr = "UPDATE db_brand.dbo.tbl_2013brand_image SET" + vbcrlf
			sqlStr = sqlStr & " image_order = "&tmpSort&"" + vbcrlf
			sqlStr = sqlStr & " WHERE idx =" + fidx(i)
			dbget.execute sqlStr
		End If
	Next
END IF

response.write "<script language='javascript'>"
response.write "	alert('저장되었습니다');"
If mode = "3banner" Then
	response.write "	location.replace('/admin/brand/main/index.asp?menupos="&menupos&"');"
ElseIf mode = "brandpick" Then
	response.write "	location.replace('/admin/brand/main/brandPick.asp?chgMode=2&menupos="&menupos&"');"
ElseIf mode = "interview" Then
	response.write "	location.replace('/admin/brand/main/mainInterView.asp?chgMode=3&menupos="&menupos&"');"
ElseIf mode = "lookbook" Then
	response.write "	location.replace('/admin/brand/main/mainLookBook.asp?chgMode=4&menupos="&menupos&"');"
End If
response.write "</script>"
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->