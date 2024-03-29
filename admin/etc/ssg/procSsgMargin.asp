<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim startDate, endDate, margin, isusing, strSql, idx, mode, cateCode, cateCode2, midx, delIdx, page, mallid
Dim itemid, i
startDate   = request("startDate")
endDate     = request("endDate")
margin      = request("margin")
isusing     = request("isusing")
idx         = request("idx")
mode        = request("mode")
cateCode    = request("cate")
cateCode2   = request("cate2")
midx        = request("midx")
delIdx      = request("delIdx")
itemid      = request("itemid")
page      	= request("page")
mallid      = request("mallid")

Dim cksel : cksel = request("cksel")

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

If mode = "cateDetail" Then
    If delIdx = "" Then
        strSql = ""
        strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_ssg_marginCate_detail] WHERE midx="&midx&" and cdl = '"&cateCode&"' and cdm = '"& cateCode2 &"' )"
        strSql = strSql & " BEGIN "
        strSql = strSql & "     INSERT INTO db_etcmall.[dbo].[tbl_ssg_marginCate_detail] " & vbCrLf
        strSql = strSql & "     (midx, cdl, cdm) VALUES " & vbCrLf
        strSql = strSql & "     ('"& midx &"', '"& cateCode &"', '"& cateCode2 &"') " & vbCrLf
        strSql = strSql & " END "
        dbget.Execute(strSql)
    Else
        strSql = ""
        strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_ssg_marginCate_detail] " & vbCrLf
        strSql = strSql & " WHERE idx = '"& delIdx &"' " & vbCrLf
        dbget.Execute(strSql)
    End If
    Response.Write "<script language=javascript>top.location.reload();</script>"
    dbget.close()	:	response.End
ElseIf mode = "cateMaster" then
    If idx = "" Then
        strSql = ""
        strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_marginCate_master] " & vbCrLf
        strSql = strSql & " (startDate, endDate, margin, isusing, regdate, mallid) VALUES " & vbCrLf
        strSql = strSql & " ('"& startDate &"', '"& endDate &"', '"& margin &"', '"& isusing &"', getdate(), '"& mallid &"') "
        dbget.Execute(strSql)
    Else
        strSql = ""
        strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ssg_marginCate_master] SET " & vbCrLf
        strSql = strSql & " startDate = '"& startDate &"'" & vbCrLf
        strSql = strSql & " ,endDate = '"& endDate &"'" & vbCrLf
        strSql = strSql & " ,margin = '"& margin &"'" & vbCrLf
        strSql = strSql & " ,isusing = '"& isusing &"'" & vbCrLf
        strSql = strSql & " WHERE idx = '"& idx &"' "
        strSql = strSql & " and mallid = '"& mallid &"' "
        dbget.Execute(strSql)
    End If
    Response.Write "<script language=javascript>alert('저장 하였습니다.');top.location.reload();</script>"
    dbget.close()	:	response.End
ElseIf mode = "itemDetail" Then
    If delIdx = "" Then
		itemid = Split(itemid, ",")
		for i = 0 to UBound(itemid)
			if Trim(itemid(i)) <> "" then
				strSql = ""
				strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_ssg_marginItem_detail] WHERE midx="&midx&" and itemid = '"&itemid(i)&"' )"
				strSql = strSql & " BEGIN "
				strSql = strSql & "     INSERT INTO db_etcmall.[dbo].[tbl_ssg_marginItem_detail] " & vbCrLf
				strSql = strSql & "     (midx, itemid) VALUES " & vbCrLf
				strSql = strSql & "     ('"& midx &"', '"& itemid(i) &"') " & vbCrLf
				strSql = strSql & " END "
				''response.write strSql
				dbget.Execute(strSql)
			end if
		next
		page = "1"
    Else
        strSql = ""
        strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_ssg_marginItem_detail] " & vbCrLf
        strSql = strSql & " WHERE idx = '"& delIdx &"' " & vbCrLf
        dbget.Execute(strSql)
    End If
    Response.Write "<script language=javascript>location.href='popSsgMarginItemDetail.asp?page=" & page & "&midx=" & midx & "&mallid=" & mallid & "'</script>"
    dbget.close()	:	response.End
ElseIf mode = "itemMaster" then
    If idx = "" Then
        strSql = ""
        strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_marginItem_master] " & vbCrLf
        strSql = strSql & " (startDate, endDate, margin, isusing, regdate, mallid) VALUES " & vbCrLf
        strSql = strSql & " ('"& startDate &"', '"& endDate &"', '"& margin &"', '"& isusing &"', getdate(), '"&mallid&"') "
        dbget.Execute(strSql)
    Else
        strSql = ""
        strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_ssg_marginItem_master] SET " & vbCrLf
        strSql = strSql & " startDate = '"& startDate &"'" & vbCrLf
        strSql = strSql & " ,endDate = '"& endDate &"'" & vbCrLf
        strSql = strSql & " ,margin = '"& margin &"'" & vbCrLf
        strSql = strSql & " ,isusing = '"& isusing &"'" & vbCrLf
        strSql = strSql & " WHERE idx = '"& idx &"' "
        strSql = strSql & " and mallid = '"& mallid &"' "
        dbget.Execute(strSql)
    End If
    Response.Write "<script language=javascript>alert('저장 하였습니다.');top.location.reload();</script>"
    dbget.close()	:	response.End
ElseIf mode = "selDel" Then
	cksel = Trim(cksel)
	If Right(cksel,1) = "," Then cksel = Left(cksel, Len(cksel) - 1)

	strSql = ""
	strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_ssg_marginItem_detail] "	 & vbcrlf
	strSql = strSql & " WHERE idx in (" & cksel & ")" & VbCrlf
	dbget.Execute strSql
    Response.Write "<script language=javascript>location.href='popSsgMarginItemDetail.asp?page=" & page & "&midx=" & midx & "&mallid=" & mallid & "'</script>"
    dbget.close()	:	response.End
else
	Response.Write "<script language=javascript>alert('잘못된 접근입니다.');</script>"
	Response.Write "잘못된 접근입니다."
    dbget.close()	:	response.End
End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
