<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg
mode = Request("mode")

'// 상품번호/옵션번호를 받는다 //
Dim dispNo, cdl, cdm, cds, infodiv, CdmKey, safecode, isvat, brandcd, makerid, mdid, catekey
dispNo	= requestCheckvar(Request("dspNo"),32)
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
CdmKey	= requestCheckvar(Request("CdmKey"),10)
safecode= requestCheckvar(Request("safecode"),10)
isvat	= requestCheckvar(Request("isvat"),10)
infodiv	= requestCheckvar(Request("infodiv"),10)
brandcd	= Request("brandcd")
makerid = requestCheckvar(Request("makerid"),32)
mdid	= Request("mdid")
catekey = Request("catekey")

If (mode = "saveCate") OR (mode = "delGbn") OR (mode = "delPrddiv") Then
	If (dispNo = "" ) OR cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
        '중복 확인
        sqlStr = "Select count(*) as cnt From db_item.dbo.tbl_gsshop_prdDiv_mapping "  & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and divcode='" & dispNo & "'"
		rsget.Open sqlStr,dbget,1
		If rsget("cnt") = 0 Then
			'신규등록
			sqlStr = ""
			sqlStr = sqlStr & " Insert into db_item.dbo.tbl_gsshop_prdDiv_mapping  " & VbCrlf
			sqlStr = sqlStr & " (divcode, infodiv, tenCateLarge, tenCateMid, tenCateSmall, safecode, isvat, lastUpdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
			sqlStr = sqlStr & ", '"&infodiv&"', '" & cdl & "','" & cdm & "','" & cds & "', '"& safecode &"','"& isvat &"', getdate()) "
			dbget.execute sqlStr
		Else
		    iErrMsg = "이미 매핑된 상품분류는 ["&dispNo&"] 추가할 수 없습니다."
		End If
		rsget.Close
	Case "saveMD"
        '중복 확인
        sqlStr = "Select CateKey From db_item.dbo.tbl_gsshop_mdid_mapping "  & VbCrlf
		sqlStr = sqlStr& " WHERE CateKey='" & catekey & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and mdid='" & mdid & "'"
		rsget.Open sqlStr,dbget,1
		If rsget.EOF Then
			'신규등록
			sqlStr = ""
			sqlStr = sqlStr & " Insert into db_item.dbo.tbl_gsshop_mdid_mapping  " & VbCrlf
			sqlStr = sqlStr & " (CateKey, mdid, regdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & CateKey & "', '"&mdid&"', getdate())"
			dbget.execute sqlStr
		Else
		    iErrMsg = "이미 매핑된 MDID는 ["&CateKey&"] 추가할 수 없습니다."
		End If
		rsget.Close
	Case "savebrandcd"
        '중복 확인
        sqlStr = "Select Count(*) as cnt From db_item.dbo.tbl_gsshop_brandcd_mapping "  & VbCrlf
		sqlStr = sqlStr& " Where makerid='" & makerid & "'"  & VbCrlf
		rsget.Open sqlStr,dbget,1
		If rsget("cnt") > 0 Then
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE db_item.dbo.tbl_gsshop_brandcd_mapping SET " & VbCrlf
			sqlStr = sqlStr & " brandcd = '"& brandcd &"' " 
			sqlStr = sqlStr & " ,lastUpdate = getdate() " 
			sqlStr = sqlStr & " ,updateid = '"& session("ssBctID") &"' " 
			sqlStr = sqlStr & " WHERE makerid = '"& makerid &"' " 
			dbget.execute sqlStr
		Else
			sqlStr = ""
			sqlStr = sqlStr & " Insert into db_item.dbo.tbl_gsshop_brandcd_mapping  " & VbCrlf
			sqlStr = sqlStr & " (makerid, brandcd, regdate, regid) " & VbCrlf
			sqlStr = sqlStr & " VALUES('"& makerid &"', '" & brandcd & "', getdate(), '"& session("ssBctID") &"')"
			dbget.execute sqlStr
		End If
		rsget.Close
	Case "delbrandcd"
		'매칭된 텐바이텐 브랜드 삭제
		sqlStr = "Delete From db_item.dbo.tbl_gsshop_brandcd_mapping " & VbCrlf
		sqlStr = sqlStr& " Where makerid='" & makerid & "'" & VbCrlf
		dbget.execute(sqlStr)
	Case "delPrddiv"
		'매칭된 텐바이텐 카테고리 삭제
		sqlStr = "Delete From db_item.dbo.tbl_gsshop_prdDiv_mapping " & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and divcode='" & dispNo & "'"
		sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'"
		dbget.execute(sqlStr)
	Case "delMdid"
		sqlStr = "Delete From db_item.dbo.tbl_gsshop_mdid_mapping " & VbCrlf
		sqlStr = sqlStr& " Where mdid='" & mdid & "'" & VbCrlf
		sqlStr = sqlStr& " 	and Catekey='" & catekey & "'" & VbCrlf
		dbget.execute(sqlStr)
End Select

%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->