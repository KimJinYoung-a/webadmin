<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/gmarket/gmarketCls.asp"-->
<!-- #include virtual="/admin/etc/gmarket/incgmarketFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, failCnt, chgSellYn, goodsGrpCd, cateCode, mallInfoDiv
Dim resultMessage, strSql, AssignedRow, ccd, selectCateCode, arrRows, i, depth
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
goodsGrpCd		= request("goodsGrpCd")
ccd				= request("ccd")
depth			= request("depth")
failCnt			= 0

If action = "ebayCommonCode" Then
	If ccd = "brand" or ccd = "maker" or ccd = "placepolicy" or ccd = "infocodedtl" or ccd = "mastercode" or ccd = "sitecode" or ccd = "addon" Then
		Call fnebayCommonCode(ccd, goodsGrpCd)
		response.end
	ElseIf ccd = "dispcategory" Then
		'webadmin에서 호출 뺐음..강제로 URL 호출해서 끌어오자.
		''0.	GUBUN : auction1010 or gmarket1010으로 바꾸기, DELETE FROM db_temp.dbo.tbl_ebay_siteCategory 해서 비우기
		''1.	http://localhost:11117/admin/etc/gmarket/gmarketActProc.asp?act=ebayCommonCode&ccd=dispcategory&depth=1 호출 1depth처리
		''2.	http://localhost:11117/admin/etc/gmarket/gmarketActProc.asp?act=ebayCommonCode&ccd=dispcategory&depth=2 호출 2depth처리
		''3.	http://localhost:11117/admin/etc/gmarket/gmarketActProc.asp?act=ebayCommonCode&ccd=dispcategory&depth=3 호출 3depth처리
		''4.	http://localhost:11117/admin/etc/gmarket/gmarketActProc.asp?act=ebayCommonCode&ccd=dispcategory&depth=4 호출 4depth처리
		''5.	http://localhost:11117/admin/etc/gmarket/gmarketActProc.asp?act=ebayCommonCode&ccd=dispcategory&depth=o 호출 재귀함수로 실제 테이블 만들기
		If depth = "1" Then
			strSql = ""
			strSql = strSql & " DELETE FROM db_temp.dbo.tbl_ebay_siteCategory WHERE gubun = '"& CMALLNAME &"' and depth = '"& depth &"' "
			dbget.execute(strSql)
			Call fnebaySiteCategoryCode(depth, cateCode, CMALLNAME)
		ElseIf depth = "o" Then
			strSql = ""
			strSql = strSql & " DELETE FROM db_etcmall.dbo.tbl_ebay_siteCategory WHERE gubun = '"& CMALLNAME &"' "
			dbget.execute(strSql)

			strSql = ""
			strSql = strSql & " ;WITH CTETABLE(catcode, parentCatcode, catname, catname2, LV, isLeaf) as ( "
			strSql = strSql & " 	SELECT A.catcode, A.parentCatcode "
			strSql = strSql & " 	, convert(varchar(300), A.catname) as catname "
			strSql = strSql & " 	, catname as catname2 "
			strSql = strSql & " 	, 1 "
			strSql = strSql & " 	, A.isLeaf "
			strSql = strSql & " 	FROM db_temp.dbo.tbl_ebay_siteCategory  A with(nolock)"
			strSql = strSql & " 	WHERE A.parentCatcode = '0' "
			strSql = strSql & " 	AND A.gubun = '"& CMALLNAME &"' "
			strSql = strSql & " 	UNION ALL "
			strSql = strSql & " 	SELECT B.catcode, B.parentCatcode "
			strSql = strSql & " 	, convert(varchar(300), C.catname + ' > ' + B.catName) as catname "
			strSql = strSql & " 	, B.catName as catname2 "
			strSql = strSql & " 	, (C.LV + 1) LV "
			strSql = strSql & " 	, B.isLeaf "
			strSql = strSql & " 	FROM db_temp.dbo.tbl_ebay_siteCategory  B with(nolock), "
			strSql = strSql & " 	CTETABLE C "
			strSql = strSql & " 	WHERE B.parentCatcode = C.catcode "
			strSql = strSql & " 	AND B.gubun = '"& CMALLNAME &"' "
			strSql = strSql & " ) "

			strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_ebay_siteCategory (cateCode, parentCateCode, cateName, cateName2, LV, regdate, gubun) "
			strSql = strSql & " SELECT catcode, parentCatcode, catname, catname2, LV, getdate(), '"& CMALLNAME &"' "
			strSql = strSql & " FROM CTETABLE "
			strSql = strSql & " WHERE isLeaf = 'Y' "
			strSql = strSql & " GROUP BY catcode, parentCatcode, catname, catname2, LV "
			strSql = strSql & " ORDER BY catname, LV "
			dbget.execute(strSql)
			rw "ok"
		Else
			strSql = ""
			strSql = strSql & " SELECT catCode "
			strSql = strSql & " FROM db_temp.dbo.tbl_ebay_siteCategory "
			strSql = strSql & " WHERE depth = '"&depth - 1&"' "
			strSql = strSql & " AND gubun = '"&CMALLNAME&"' "
			strSql = strSql & " AND isLeaf = 'N' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If isArray(arrRows) Then
				strSql = ""
				strSql = strSql & " DELETE FROM db_temp.dbo.tbl_ebay_siteCategory WHERE gubun = '"& CMALLNAME &"' and depth = '"& depth &"' "
				dbget.execute(strSql)
				For i = 0 To UBound(arrRows,2)
					Call fnebaySiteCategoryCode(depth, arrRows(0, i), CMALLNAME)
					If (i mod 50) = 0 Then
						rw "호출중 : " & i
						response.flush
						response.clear
					End If
				Next
			End If
			rw "END"
		End If
		response.end
	ElseIf ccd = "matchcategory" Then
		strSql = ""
		strSql = strSql & " SELECT SDCategoryCode "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_ebay_esmCategory "
		strSql = strSql & " ORDER BY SDCategoryCode DESC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		If isArray(arrRows) Then
			For i = 0 To UBound(arrRows,2)
				Call fnebayCommonCode(ccd, arrRows(0, i))
				If (i mod 50) = 0 Then
					rw "호출중 : " & i
					response.flush
					response.clear
				End If
			Next
		End If
		rw "END"
	ElseIf ccd = "editNamebysitecategory" Then
		strSql = ""
		strSql = strSql & " SELECT cateCode "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_ebay_siteCategory "
		strSql = strSql & " WHERE gubun = '"&CMALLNAME&"' "
		strSql = strSql & " GROUP BY cateCode "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		If isArray(arrRows) Then
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ebay_siteCategory "
			strSql = strSql & " SET isEditName = NULL "
			strSql = strSql & " WHERE gubun = '"&CMALLNAME&"' "
			dbget.execute(strSql)
			For i = 0 To UBound(arrRows,2)
				Call fnebayCommonCode(ccd, arrRows(0, i))
				If (i mod 50) = 0 Then
					rw "호출중 : " & i
					response.flush
					response.clear
				End If
			Next
		End If
		rw "END"
	ElseIf ccd = "optionPolicy" Then
		strSql = ""
		strSql = strSql & " SELECT cateCode "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_ebay_siteCategory "
		strSql = strSql & " WHERE gubun = '"&CMALLNAME&"' "
		strSql = strSql & " GROUP BY cateCode "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		If isArray(arrRows) Then
			' strSql = ""
			' strSql = strSql & " UPDATE db_etcmall.dbo.tbl_ebay_siteCategory "
			' strSql = strSql & " SET optionUseYn = NULL, customOptionUseYn = NULL "
			' strSql = strSql & " ,genrlUseYn = NULL, twoCmbtUseYn = NULL "
			' strSql = strSql & " ,threeCmbtUseYn = NULL, textUseYn = NULL "
			' strSql = strSql & " ,calcUseYn = NULL, addAmntUseYn = NULL "
			' strSql = strSql & " WHERE gubun = '"&CMALLNAME&"' "
			' dbget.execute(strSql)
			For i = 0 To UBound(arrRows,2)
				Call fnebayCommonCode(ccd, arrRows(0, i))
				If (i mod 50) = 0 Then
					rw "호출중 : " & i
					response.flush
					response.clear
				End If
			Next
		End If
		rw "END"
	ElseIf ccd = "rcmdOption" Then
		strSql = ""
		strSql = strSql & " SELECT cateCode "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_ebay_siteCategory "
		strSql = strSql & " WHERE gubun = '"&CMALLNAME&"' "
		strSql = strSql & " and isUseTenOpt is null "
		strSql = strSql & " GROUP BY cateCode "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		If isArray(arrRows) Then
			For i = 0 To UBound(arrRows,2)
				Call fnebayCommonCode(ccd, arrRows(0, i))
				If (i mod 50) = 0 Then
					rw "호출중 : " & i
					response.flush
					response.clear
				End If
			Next
		End If
		rw "END"
	Else
		Call fnebayCommonCode(ccd, "")
		response.end
	End If
' ElseIf action = "updateSendState" Then						'주문상태변경 / benepia_SongjangProc.asp에서 넘어온다.
' 	AssignedRow = fnbenepiaSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
' 	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
' 	response.end
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&resultMessage&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->