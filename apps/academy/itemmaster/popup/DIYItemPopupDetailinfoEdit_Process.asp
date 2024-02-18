<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<%
dim i,k

dim refer
''// 유효 접근 주소 검사 //
refer = request.ServerVariables("HTTP_REFERER")

Dim itemid, requirecontents, keywords, ordercomment, mode
Dim sqlStr, ArrCnt, makerid

itemid = requestCheckvar(Request("itemid"),10)
requirecontents = html2db(Request.form("requirecontents"))
keywords = html2db(Request.form("keywords"))
ordercomment = html2db(Request.form("ordercomment"))
mode = Request.form("mode")
makerid = request.cookies("partner")("userid")

If (ItemCheckMyItemYN(makerid,itemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If

dim foundcount, found, ArroptionName, Arroptlimitno
Dim limityn, Arroptaddprice, Arroptaddbuyprice, Arroptioncode
If itemid <> "" Then
	If mode = "cate" Then
		'###########################################################################
		'전시카테고리 넣기
		'###########################################################################
		Dim  chkCateDef,CateCode, CateDepth, isDefault
		chkCateDef = 0
		CateCode = Split(Request.Form("arrcatecode"),",")
		CateDepth = Split(Request.Form("arrcatedepth"),",")
		isDefault = Split(Request.Form("arrisdefault"),",")
		ArrCnt = ubound(CateCode)

		If (Request.Form("arrcatecode")<>"") Then 
			sqlStr = "delete from db_academy.dbo.tbl_display_cate_item_Academy Where itemid='" & CStr(itemid) & "';" & vbCrLf
			dbACADEMYget.execute(sqlStr)
			for i=0 to ArrCnt
				'2015.06.18 수정 (기본카테고리는 하나만 설정되게)
				if UCase(isDefault(i)) ="Y" and chkCateDef = 1 then
					isDefault(i)="N"
				end if
				sqlStr = "Insert into db_academy.dbo.tbl_display_cate_item_Academy (catecode, itemid, depth, sortNo, isDefault) values "
				sqlStr = sqlStr & "('" & Cstr(CateCode(i)) & "'"
				sqlStr = sqlStr & ",'" & CStr(itemid) & "'"
				sqlStr = sqlStr & ",'" & CStr(CateDepth(i)) & "',9999"
				sqlStr = sqlStr & ",'" & CStr(isDefault(i)) & "');" & vbCrLf
				dbACADEMYget.execute(sqlStr)
				'Response.write sqlStr &"<br>"
				IF UCase(isDefault(i)) ="Y" THEN '기본 카테고리 설정되어 있는지 확인
					chkCateDef = 1
				END IF
			Next
		end If
		'###########################################################################
	ElseIf mode = "iteminfo" Then
		'###########################################################################
		'상품 품목고시정보 저장 
		'###########################################################################
		if Request.Form("infoDiv")<>"" then
			dim infoCd, infoCont, infoChk
			'배열로 처리
			redim infoCd(Request.Form("infoCd").Count)
			redim infoCont(Request.Form("infoCont").Count)
			redim infoChk(Request.Form("infoChk").Count)
			for i=1 to Request.Form("infoCd").Count
				infoCd(i) = Request.Form("infoCd")(i)
				infoCont(i) = Request.Form("infoCont")(i)
				infoChk(i) = Request.Form("infoChk")(i)
			next

			'기존값 삭제
			sqlStr = "Delete From db_academy.dbo.tbl_diy_item_infoCont Where itemid='" & CStr(itemid) & "'"
			dbACADEMYget.execute(sqlStr)

			'DB에 처리
			for i=1 to ubound(infoCd)
				'입력값이 있는 경우만 저장
				if infoChk(i)<>"" or infoCont(i)<>"" then
					sqlStr = "Insert into db_academy.dbo.tbl_diy_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
					sqlStr = sqlStr & "('" & CStr(itemid) & "'"
					sqlStr = sqlStr & ",'" & CStr(Trim((infoCd(i)))) & "'"
					sqlStr = sqlStr & ",'" & CStr(Trim(infoChk(i))) & "'"
					sqlStr = sqlStr & ",'" & html2db(Trim(infoCont(i))) & "')"
					'Response.write sqlStr & "<br>"
					dbACADEMYget.execute(sqlStr)
				end if
			Next
			sqlStr = "update db_academy.dbo.tbl_diy_item_Contents" + vbCrlf
			sqlStr = sqlStr & " set infoDiv='" & Request.Form("infoDiv") & "'" + vbCrlf
			sqlStr = sqlStr & " where itemid=" + CStr(itemid) + vbCrlf
			dbACADEMYget.Execute sqlStr
		End If
	ElseIf mode = "voddel" Then
		'###########################################################################
		'동영상 삭제 
		'###########################################################################
		sqlStr = "Delete From db_academy.dbo.tbl_diy_item_videos WHERE ITEMID='" & CStr(itemid) & "'"
		dbACADEMYget.execute(sqlStr)
	Else
		'###########################################################################
		'상품 데이터 수정
		'###########################################################################

		sqlStr = "update db_academy.dbo.tbl_diy_item_Contents" + vbCrlf
		If requirecontents <> "" Then
		sqlStr = sqlStr & " set requirecontents='" & requirecontents & "'" + vbCrlf
		End If
		If keywords <> "" Then
		sqlStr = sqlStr & " set keywords='" & keywords & "'" + vbCrlf
		End If
		If ordercomment <> "" Then
		sqlStr = sqlStr & " set ordercomment='" & ordercomment & "'" + vbCrlf
		End If
		If Request.form("safetyYn") <> "" Then
		sqlStr = sqlStr & " set safetyYn='" & Request.form("safetyYn") & "'" + vbCrlf
		sqlStr = sqlStr & " ,safetyDiv='" & Request.form("safetyDiv") & "'" + vbCrlf
		sqlStr = sqlStr & " ,safetyNum='" & html2db(Request.form("safetyNum")) & "'" + vbCrlf
		End If
		sqlStr = sqlStr & " where itemid=" + CStr(itemid) + vbCrlf
		dbACADEMYget.Execute sqlStr
		'###########################################################################
	End If
End If
%>
<script>
<!--
	parent.fnDetailInfoEnd();
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->