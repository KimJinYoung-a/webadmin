<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg, categbn, makerid, BrandCode


mode    = Request("mode")
categbn = Request("categbn")
''If mode <> "saveAddress" Then
''	If (categbn <> "cate") AND (categbn <> "brand") Then
''		response.write "<script>alert('잘못된 경로입니다');window.close();</script>"
''		response.end
''	End If
''End If

Dim cdl, cdm, cds, depthCode, stdcode, i
Dim chkArr, depthCodeArr, stdcodeArr, sitenoArr
'makerid	= requestCheckvar(Request("makerid"),32)
'BrandCode = requestCheckvar(Request("BrandCode"),32)

cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
depthCode = requestCheckvar(Request("depthCode"),16)
stdcode   = requestCheckvar(Request("stdcode"),16)


If (mode = "saveCate") or (mode="saveCateArr") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
    CASE "saveCateArr"
        set chkArr = Request.form("chk")
        set stdcodeArr = Request.form("stdcode")
        set depthCodeArr= Request.form("depthcode")
        set sitenoArr  = Request.form("siteno")
        
       
        for i=1 to chkArr.count
            sqlStr = "exec db_etcmall.dbo.usp_Ten_outmall_Ssg_categoryMapp '"&cdl&"','"&cdm&"','"&cds&"','"&stdcodeArr(chkArr(i)+1)&"','"&depthCodeArr(chkArr(i)+1)&"','"&sitenoArr(chkArr(i)+1)&"'"
            dbget.execute(sqlStr)
        next
            
        set chkArr = Nothing
        set stdcodeArr = Nothing
        set depthCodeArr = Nothing
        set sitenoArr = Nothing
        
	Case "saveCate"
        '중복 확인
'        If categbn = "cate" Then
''			sqlStr = "DELETE FROM db_etcmall.dbo.tbl_ssg_cate_mapping " & VbCrlf
''			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
''			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
''			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
''			dbget.execute(sqlStr)
'        	
'	        sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_ssg_cate_mapping "  & VbCrlf
'			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
'			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
'			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
'			sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
'		    sqlStr = sqlStr& " 	and stdCtgDClsCd='" & stdcode & "'"
'			rsget.Open sqlStr,dbget,1
'			If rsget("cnt") > 0 Then
'			     joongBok = True
'			End If
'			rsget.Close
'		End If
'
'		If joongBok = False Then
'			sqlStr = ""
'			sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_ssg_cate_mapping  " & VbCrlf
'			sqlStr = sqlStr & " (stdCtgDClsCd, depthCode, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
'			sqlStr = sqlStr & " VALUES('"&stdcode&"','" & depthCode & "' "  & VbCrlf
'			sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
'			dbget.execute sqlStr
'		Else
'		    iErrMsg = "이미 매핑된 카테고리는  추가할 수 없습니다."
'		End If
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_ssg_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
		sqlStr = sqlStr& " 	and stdCtgDClsCd='" & stdcode & "'"
		dbget.execute(sqlStr)
'	Case "saveBrand"
'        '중복 확인
'        If categbn = "brand" Then
'	        sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_ssg_brand_mapping "  & VbCrlf
'			sqlStr = sqlStr& " WHERE makerid='" & makerid & "'"  & VbCrlf
'			rsget.Open sqlStr,dbget,1
'			If rsget("cnt") > 0 Then
'			     joongBok = True
'			End If
'			rsget.Close
'		End If
'
'		If joongBok = False Then
'			sqlStr = ""
'			sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_ssg_brand_mapping  " & VbCrlf
'			sqlStr = sqlStr & " (makerid, BrandCode)" & VbCrlf
'			sqlStr = sqlStr & " VALUES('" & makerid & "' "  & VbCrlf
'			sqlStr = sqlStr & ", '" & BrandCode & "') "
'			dbget.execute sqlStr
'		Else
'		    iErrMsg = "이미 매핑된 브랜드는 추가할 수 없습니다."
'		End If
'	Case "delBrand"
'		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_ssg_brand_mapping " & VbCrlf
'		sqlStr = sqlStr& " WHERE BrandCode='" & BrandCode & "'" & VbCrlf
'		sqlStr = sqlStr& " and makerid='" & makerid & "'"
'		dbget.execute(sqlStr)
'
'	Case "saveAddress"
'		Dim AddressTitle, AddressName, Phone1, Phone2, reqzipcode, reqzipaddr, reqaddress
'		AddressTitle	= request("AddressTitle")
'		AddressName		= request("AddressName")
'		Phone1			= request("Phone1")
'		Phone2			= request("Phone2")
'		reqzipcode		= request("reqzipcode")
'		reqzipaddr		= request("reqzipaddr")
'		reqaddress		= request("reqaddress")
'
'		sqlStr = ""
'		sqlStr = sqlStr & " IF Exists(SELECT COUNT(*) as cnt FROM db_etcmall.[dbo].[tbl_ssg_AddressBook]) " & VbCrlf
'		sqlStr = sqlStr & " BEGIN " & VbCrlf
'		sqlStr = sqlStr & " 	UPDATE db_etcmall.[dbo].[tbl_ssg_AddressBook] SET " & VbCrlf
'		sqlStr = sqlStr & " 	AddressTitle = '"&AddressTitle&"' " & VbCrlf
'		sqlStr = sqlStr & " 	,AddressName = '"&AddressName&"' " & VbCrlf
'		sqlStr = sqlStr & " 	,Phone1 = '"&Phone1&"' " & VbCrlf
'		sqlStr = sqlStr & " 	,Phone2 = '"&Phone2&"' " & VbCrlf
'		sqlStr = sqlStr & " 	,reqzipcode = '"&reqzipcode&"' " & VbCrlf
'		sqlStr = sqlStr & " 	,reqzipaddr = '"&reqzipaddr&"' " & VbCrlf
'		sqlStr = sqlStr & " 	,reqaddress = '"&reqaddress&"' " & VbCrlf
'		sqlStr = sqlStr & " END ELSE " & VbCrlf
'		sqlStr = sqlStr & " BEGIN " & VbCrlf
'		sqlStr = sqlStr & " 	INSERT INTO db_etcmall.[dbo].[tbl_ssg_AddressBook] " & VbCrlf
'		sqlStr = sqlStr & " 	(AddressTitle, AddressName, Phone1, Phone2, reqzipcode, reqzipaddr, reqaddress) " & VbCrlf
'		sqlStr = sqlStr & " 	VALUES('" & AddressTitle & "' "  & VbCrlf
'		sqlStr = sqlStr & "		, '" & AddressName & "','" & Phone1 & "','" & Phone2 & "', '"& reqzipcode &"', '"&reqzipaddr&"', '"&reqaddress&"') " & VbCrlf
'		sqlStr = sqlStr & " END "
'		dbget.execute sqlStr
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
    alert("정상적으로 처리되었습니다.");
   
    opener.location.reload();
    parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->