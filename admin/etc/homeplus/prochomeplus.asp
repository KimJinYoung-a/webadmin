<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg, joongBok, categbn
mode = Request("mode")
categbn = Request("categbn")
If (categbn <> "dft") AND (categbn <> "disp") AND (categbn <> "branddisp") Then
	response.write "<script>alert('잘못된 경로입니다');window.close();</script>"
	response.end
End If

Dim cdl, cdm, cds, infodiv, divsioncode, groupcode, deptcode, classcode, subclasscode, categoryid, depthCode
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
infodiv	= requestCheckvar(Request("infodiv"),10)
divsioncode	= requestCheckvar(Request("divsioncode"),10)
groupcode	= requestCheckvar(Request("groupcode"),10)
deptcode	= requestCheckvar(Request("deptcode"),10)
classcode	= requestCheckvar(Request("classcode"),10)
subclasscode = requestCheckvar(Request("subclasscode"),10)
categoryid	= requestCheckvar(Request("categoryid"),10)
depthCode= requestCheckvar(Request("depthcode"),10)
joongBok = False
If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
        '중복 확인
        If categbn = "dft" Then
	        sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_homeplus_prdDiv_mapping "  & VbCrlf
			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and hDIVISION='" & divsioncode & "'"
			sqlStr = sqlStr& " 	and hGROUP='" & groupcode & "'"
			sqlStr = sqlStr& " 	and hDEPT='" & deptcode & "'"
			sqlStr = sqlStr& " 	and hCLASS='" & classcode & "'"
			sqlStr = sqlStr& " 	and hSUBCLASS='" & subclasscode & "'"
			sqlStr = sqlStr& " 	and hCATEGORY_ID='" & categoryid & "'"
			rsget.Open sqlStr,dbget,1
			If rsget("cnt") > 0 Then
			    joongBok = True
			End If
			rsget.Close
		Else
	        sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_homeplus_cate_mapping "  & VbCrlf
			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
			rsget.Open sqlStr,dbget,1
			If rsget("cnt") > 0 Then
			     joongBok = True
			End If
			rsget.Close
		End If

		If joongBok = False Then
			If categbn = "dft" Then
				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_homeplus_prdDiv_mapping  " & VbCrlf
				sqlStr = sqlStr & " (hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hCATEGORY_ID, infodiv, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
				sqlStr = sqlStr & " VALUES('" & divsioncode & "', '" & groupcode & "', '" & deptcode & "', '" & classcode & "', '" & subclasscode & "', '" & categoryid & "' "  & VbCrlf
				sqlStr = sqlStr & ", '"&infodiv&"', '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbget.execute sqlStr

				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_homeplus_prdDiv_mapping  " & VbCrlf
				sqlStr = sqlStr & " (hDIVISION, hGROUP, hDEPT, hCLASS, hSUBCLASS, hCATEGORY_ID, infodiv, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
				sqlStr = sqlStr & " VALUES('" & divsioncode & "', '" & groupcode & "', '" & deptcode & "', '" & classcode & "', '" & subclasscode & "', '" & categoryid & "' "  & VbCrlf
				sqlStr = sqlStr & ", '"&infodiv&"', '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbCTget.execute sqlStr
			Else
				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_homeplus_cate_mapping  " & VbCrlf
				sqlStr = sqlStr & " (depthCode, infodiv, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
				sqlStr = sqlStr & " VALUES('" & depthCode & "' "  & VbCrlf
				sqlStr = sqlStr & ", '"&infodiv&"', '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbget.execute sqlStr

				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_homeplus_cate_mapping  " & VbCrlf
				sqlStr = sqlStr & " (depthCode, infodiv, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
				sqlStr = sqlStr & " VALUES('" & depthCode & "' "  & VbCrlf
				sqlStr = sqlStr & ", '"&infodiv&"', '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbCTget.execute sqlStr
			End If
		Else
		    iErrMsg = "이미 매핑된 기준카테고리 및 전시카테고리는  추가할 수 없습니다."
		End If
	Case "delPrddiv"
		If categbn = "dft" Then
			sqlStr = "DELETE FROM db_etcmall.dbo.tbl_homeplus_prdDiv_mapping " & VbCrlf
			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hDIVISION='" & divsioncode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hGROUP='" & groupcode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hDEPT='" & deptcode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hCLASS='" & classcode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hSUBCLASS='" & subclasscode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hCATEGORY_ID='" & categoryid & "'"
			dbget.execute(sqlStr)

			sqlStr = "DELETE FROM db_outmall.dbo.tbl_homeplus_prdDiv_mapping " & VbCrlf
			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hDIVISION='" & divsioncode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hGROUP='" & groupcode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hDEPT='" & deptcode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hCLASS='" & classcode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hSUBCLASS='" & subclasscode & "'" & VbCrlf
			sqlStr = sqlStr& " 	and hCATEGORY_ID='" & categoryid & "'"
			dbCTget.execute(sqlStr)
		Else
			sqlStr = "DELETE FROM db_etcmall.dbo.tbl_homeplus_cate_mapping " & VbCrlf
			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'" & VbCrlf
			sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
			dbget.execute(sqlStr)

			sqlStr = "DELETE FROM db_outmall.dbo.tbl_homeplus_cate_mapping " & VbCrlf
			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'" & VbCrlf
			sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
			dbCTget.execute(sqlStr)
		End If
	Case "brandCate"
		sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_homeplus_brandcategory_mapping "  & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
		rsget.Open sqlStr,dbget,1
		If rsget("cnt") > 0 Then
			joongBok = True
		End If
		rsget.Close
		If joongBok = False Then
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_homeplus_brandcategory_mapping  " & VbCrlf
			sqlStr = sqlStr & " (depthCode, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & depthCode & "' "  & VbCrlf
			sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
			dbget.execute sqlStr

			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_homeplus_brandcategory_mapping  " & VbCrlf
			sqlStr = sqlStr & " (depthCode, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & depthCode & "' "  & VbCrlf
			sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
			dbCTget.execute sqlStr
		Else
		    iErrMsg = "이미 매핑된 카테고리는 추가 할 수 없습니다."
		End If
	Case "delbrandCate"
		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_homeplus_brandcategory_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
		dbget.execute(sqlStr)

		sqlStr = "DELETE FROM db_outmall.dbo.tbl_homeplus_brandcategory_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
		dbCTget.execute(sqlStr)
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
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->