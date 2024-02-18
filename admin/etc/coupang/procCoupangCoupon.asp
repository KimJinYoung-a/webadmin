<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim couponName, maxDiscountPrice, discount, startDate, endDate, couponType, strSql, idx, mode, cdl, cdm, midx, delIdx, page
Dim itemid, i, addSql
couponName          = request("couponName")
maxDiscountPrice    = request("maxDiscountPrice")
discount            = request("discount")
startDate           = request("startDate")
endDate             = request("endDate")
couponType          = request("couponType")
idx                 = request("idx")
mode                = request("mode")
cdl                = request("cdl")
cdm                = request("cdm")
midx                = request("midx")
delIdx              = request("delIdx")
itemid              = request("itemid")
page      	        = request("page")

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
    If Len(itemid) < 3 Then
        itemid = ""
    Else
    	itemid = left(arrItemid,len(arrItemid)-1)
    End If
End If

If mode = "cateDetail" Then
    If delIdx = "" Then
        strSql = ""
        strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_coupang_CouponCate_detail] WHERE midx="&midx&" and cdl = '"&cdl&"' and cdm = '"&cdm&"' )"
        strSql = strSql & " BEGIN "
        strSql = strSql & "     INSERT INTO db_etcmall.[dbo].[tbl_coupang_CouponCate_detail] " & vbCrLf
        strSql = strSql & "     (midx, cdl, cdm) VALUES " & vbCrLf
        strSql = strSql & "     ('"& midx &"', '"& cdl &"', '"& cdm &"') " & vbCrLf
        strSql = strSql & " END "
        dbget.Execute(strSql)
    Else
        strSql = ""
        strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_coupang_CouponCate_detail] " & vbCrLf
        strSql = strSql & " WHERE idx = '"& delIdx &"' " & vbCrLf
        dbget.Execute(strSql)
    End If
    Response.Write "<script language=javascript>top.location.reload();</script>"
    dbget.close()	:	response.End
ElseIf mode = "ItemDetail" Then
    strSql = ""
    addSql = ""
    strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_coupang_CouponItem_detail] WHERE midx="&midx&" "
    dbget.Execute(strSql)

    If (itemid <> "") then
        If Right(Trim(itemid) ,1) = "," Then
            itemid = Replace(itemid,",,",",")
            addSql = addSql & " and itemid in (" + Left(itemid,Len(itemid)-1) + ")"
        Else
            itemid = Replace(itemid,",,",",")
            addSql = addSql & " and itemid in (" + itemid + ")"
        End If
    End If

    If (itemid <> "") then
        strSql = ""
        strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_coupang_CouponItem_detail] (midx, itemid) " & vbCrLf
        strSql = strSql & " SELECT TOP 1000 '"&midx&"', itemid " & vbCrLf
        strSql = strSql & " FROM db_item.dbo.tbl_item " & vbCrLf
        strSql = strSql & " WHERE 1=1 "
        strSql = strSql & addSql
        dbget.Execute(strSql)
    End If
    Response.Write "<script language=javascript>top.location.reload();</script>"
    dbget.close()	:	response.End
ElseIf mode = "ItemDeleteDetail" Then
    strSql = ""
    addSql = ""
    strSql = strSql & " DELETE FROM db_etcmall.[dbo].[tbl_coupang_CouponItem_detail] WHERE midx="&midx&" and itemType = 'D' "
    dbget.Execute(strSql)

    If (itemid <> "") then
        If Right(Trim(itemid) ,1) = "," Then
            itemid = Replace(itemid,",,",",")
            addSql = addSql & " and itemid in (" + Left(itemid,Len(itemid)-1) + ")"
        Else
            itemid = Replace(itemid,",,",",")
            addSql = addSql & " and itemid in (" + itemid + ")"
        End If
    End If

    If (itemid <> "") then
        strSql = ""
        strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_coupang_CouponItem_detail] (midx, itemid, itemType) " & vbCrLf
        strSql = strSql & " SELECT TOP 1000 '"&midx&"', itemid, 'D' " & vbCrLf
        strSql = strSql & " FROM db_item.dbo.tbl_item " & vbCrLf
        strSql = strSql & " WHERE 1=1 "
        strSql = strSql & addSql
        dbget.Execute(strSql)
    End If
    Response.Write "<script language=javascript>top.location.reload();</script>"
    dbget.close()	:	response.End
ElseIf mode = "cateMaster" then
    If idx = "" Then
        strSql = ""
        strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_coupang_Coupon_master] " & vbCrLf
        strSql = strSql & " (couponName, maxDiscountPrice, discount, startDate, endDate, couponType, regdate) VALUES " & vbCrLf
        strSql = strSql & " ('"& couponName &"', '"& maxDiscountPrice &"', '"& discount &"', '"& startDate &"', '"& endDate &" 23:59:59', '"& couponType &"', getdate()) "
        dbget.Execute(strSql)
        Response.Write "<script language=javascript>alert('저장 하였습니다.');top.location.reload();</script>"
    Else
        strSql = ""
        strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_coupang_Coupon_master] SET " & vbCrLf
        strSql = strSql & " couponName = '"& couponName &"'" & vbCrLf
        strSql = strSql & " ,maxDiscountPrice = '"& maxDiscountPrice &"'" & vbCrLf
        strSql = strSql & " ,discount = '"& discount &"'" & vbCrLf
        strSql = strSql & " ,startDate = '"& startDate &"'" & vbCrLf
        strSql = strSql & " ,endDate = '"& endDate &" 23:59:59'" & vbCrLf
        strSql = strSql & " ,couponType = '"& couponType &"'" & vbCrLf
        strSql = strSql & " WHERE idx = '"& idx &"' "
        dbget.Execute(strSql)
        Response.Write "<script language=javascript>alert('저장 하였습니다.');top.location.replace('/admin/etc/coupang/popCoupangCouponCateList.asp');</script>"
    End If
    dbget.close()	:	response.End
else
	Response.Write "<script language=javascript>alert('잘못된 접근입니다.');</script>"
	Response.Write "잘못된 접근입니다."
    dbget.close()	:	response.End
End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
