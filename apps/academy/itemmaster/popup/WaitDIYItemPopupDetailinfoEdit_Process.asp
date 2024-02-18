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
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<%
dim i,k

dim refer
''// 유효 접근 주소 검사 //
refer = request.ServerVariables("HTTP_REFERER")

Dim waititemid, requirecontents, keywords, ordercomment, mode
Dim sqlStr, ArrCnt, makerid
dim infoCd, infoCont, infoChk
dim DesignerID

waititemid = requestCheckvar(Request("waititemid"),10)
requirecontents = html2db(Request.form("requirecontents"))
keywords = html2db(Request.form("keywords"))
ordercomment = html2db(Request.form("ordercomment"))
mode = Request.form("mode")
makerid = request.cookies("partner")("userid")

dim foundcount, found
Dim limityn, Arroptioncode, optlimitno

If waititemid = "" And mode = "iteminfofirst" Then
	DesignerID = Request.Form("designerid")
	'###########################################################################
	'상품 데이터 입력
	'###########################################################################
	sqlStr = "insert into db_academy.dbo.tbl_diy_wait_item" + vbCrlf
	sqlStr = sqlStr & " (itemdiv,makerid,itemname,regdate,buycash, sellcash, mileage, sellyn, deliverytype,limityn,currstate)" + vbCrlf
	sqlStr = sqlStr & " values(" + vbCrlf
	sqlStr = sqlStr & "'01'" + vbCrlf
	sqlStr = sqlStr & ",'" & DesignerID & "'" + vbCrlf
	sqlStr = sqlStr & ",'tempitem'" + vbCrlf
	sqlStr = sqlStr & ",getdate()" + vbCrlf
	sqlStr = sqlStr & ",0" + vbCrlf
	sqlStr = sqlStr & ",0" + vbCrlf
	sqlStr = sqlStr & ",0" + vbCrlf
	sqlStr = sqlStr & ",'N'" + vbCrlf
	sqlStr = sqlStr & ",'9'" + vbCrlf
	sqlStr = sqlStr & ",'N'" + vbCrlf
	sqlStr = sqlStr & ",3)" + vbCrlf
	'Response.write sqlStr
	'Response.end
	dbACADEMYget.Execute sqlStr
	'###########################################################################
	'상품 아이디 가져오기
	'###########################################################################
	sqlStr = "Select IDENT_CURRENT('db_academy.dbo.tbl_diy_wait_item') as maxitemid "
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
		waititemid = rsACADEMYget("maxitemid")
	rsACADEMYget.close
End If

If waititemid <> "" Then
	If mode = "cate" Then
		If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
			Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
			Response.end
		End If
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
			sqlStr = "delete from db_academy.dbo.tbl_display_cate_waitItem_Academy Where itemid='" & CStr(waititemid) & "';" & vbCrLf
			dbACADEMYget.execute(sqlStr)
			for i=0 to ArrCnt
				'2015.06.18 수정 (기본카테고리는 하나만 설정되게)
				if UCase(isDefault(i)) ="Y" and chkCateDef = 1 then
					isDefault(i)="N"
				end if
				sqlStr = "Insert into db_academy.dbo.tbl_display_cate_waitItem_Academy (catecode, itemid, depth, sortNo, isDefault) values "
				sqlStr = sqlStr & "('" & Cstr(CateCode(i)) & "'"
				sqlStr = sqlStr & ",'" & CStr(waititemid) & "'"
				sqlStr = sqlStr & ",'" & CStr(CateDepth(i)) & "',9999"
				sqlStr = sqlStr & ",'" & CStr(isDefault(i)) & "');" & vbCrLf
				dbACADEMYget.execute(sqlStr)
				'Response.write sqlStr &"<br>"
				IF UCase(isDefault(i)) ="Y" THEN '기본 카테고리 설정되어 있는지 확인
					chkCateDef = 1
				END IF
			Next
		end If
	ElseIf mode = "iteminfo" Then
		'###########################################################################
		'상품 품목고시정보 저장 
		'###########################################################################
		If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
			Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
			Response.end
		End If
		if Request.Form("infoDiv")<>"" then
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
			sqlStr = "Delete From db_academy.dbo.tbl_diy_wait_item_infoCont Where itemid='" & CStr(waititemid) & "'"
			dbACADEMYget.execute(sqlStr)

			'DB에 처리
			for i=1 to ubound(infoCd)
				'입력값이 있는 경우만 저장
				if infoChk(i)<>"" or infoCont(i)<>"" then
					sqlStr = "Insert into db_academy.dbo.tbl_diy_wait_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
					sqlStr = sqlStr & "('" & CStr(waititemid) & "'"
					sqlStr = sqlStr & ",'" & CStr(infoCd(i)) & "'"
					sqlStr = sqlStr & ",'" & CStr(infoChk(i)) & "'"
					sqlStr = sqlStr & ",'" & html2db(infoCont(i)) & "')"
					dbACADEMYget.execute(sqlStr)
				end if
			Next
			sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + vbCrlf
			sqlStr = sqlStr & " set infoDiv='" & Request.Form("infoDiv") & "'" + vbCrlf
			sqlStr = sqlStr & " where itemid=" + CStr(waititemid) + vbCrlf
			dbACADEMYget.Execute sqlStr
		End If
	ElseIf mode = "iteminfofirst" Then
		'###########################################################################
		'상품 품목고시정보 저장 
		'###########################################################################
		if Request.Form("infoDiv")<>"" then
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
			sqlStr = "Delete From db_academy.dbo.tbl_diy_wait_item_infoCont Where itemid='" & CStr(waititemid) & "'"
			dbACADEMYget.execute(sqlStr)

			'DB에 처리
			for i=1 to ubound(infoCd)
				'입력값이 있는 경우만 저장
				if infoChk(i)<>"" or infoCont(i)<>"" then
					sqlStr = "Insert into db_academy.dbo.tbl_diy_wait_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
					sqlStr = sqlStr & "('" & CStr(waititemid) & "'"
					sqlStr = sqlStr & ",'" & CStr(infoCd(i)) & "'"
					sqlStr = sqlStr & ",'" & CStr(infoChk(i)) & "'"
					sqlStr = sqlStr & ",'" & html2db(infoCont(i)) & "')"
					dbACADEMYget.execute(sqlStr)
				end if
			Next
			sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + vbCrlf
			sqlStr = sqlStr & " set infoDiv='" & Request.Form("infoDiv") & "'" + vbCrlf
			sqlStr = sqlStr & " where itemid=" + CStr(waititemid) + vbCrlf
			dbACADEMYget.Execute sqlStr
		End If
	ElseIf mode = "voddel" Then
		If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
			Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
			Response.end
		End If
		'###########################################################################
		'동영상 삭제 
		'###########################################################################
		sqlStr = "Delete From db_academy.dbo.tbl_diy_wait_item_videos where itemid='" & CStr(waititemid) & "'"
		dbACADEMYget.execute(sqlStr)
	ElseIf mode = "editOption" Or mode = "editOptionMultiple" Then
		If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
			Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
			Response.end
		End If
		'###########################################################################
		'상품 옵션 넣기
		'###########################################################################
'Response.write Request.Form("optionnameset")
'Response.end

		redim ArroptionName(Request.Form("optionName").Count)
		redim Arroptaddprice(Request.Form("optaddprice").Count)
		redim Arroptaddbuyprice(Request.Form("optaddbuyprice").Count)
		
		optlimitno=0
		Arroptioncode = Split(Request.Form("optioncode"),",")
		ArrCnt = Request.Form("optionName").Count

		limityn = "N"
		if Request.Form("useoptionyn")="Y" then
			''단일 옵션
			if (Request.Form("optlevel")="1") Then
				If ArrCnt > 0 Then
					'################### 기존 등록 옵션 삭제 #######################################################
					sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option"
					sqlStr = sqlStr & " where itemid=" & waititemid
					rsACADEMYget.Open sqlStr,dbACADEMYget,1
					'################### 기존 등록 옵션 삭제 #######################################################
					sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option_Multiple"
					sqlStr = sqlStr & " where itemid=" & waititemid
					rsACADEMYget.Open sqlStr,dbACADEMYget,1
				End If

				For i=1 To ArrCnt
					ArroptionName(i) = Request.Form("optionName")(i)
					Arroptaddprice(i) = Request.Form("optaddprice")(i)
					Arroptaddbuyprice(i) = Request.Form("optaddbuyprice")(i)
					If optlimitno > 0 Then
						limityn = "Y"
					End If
					if (Trim(Arroptioncode(i-1)) <> "") then
						''중복 옵션은 안올림. 한정 구분은 상품 한정 구분과 동일
							sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold , optaddPrice, optaddBuyPrice)"
							sqlStr = sqlStr + " values(" + CStr(waititemid) + ", '" + CStr(Trim(Arroptioncode(i-1))) + "', convert(varchar(32),'" + html2db(Request.Form("optionTypename")) + "'), convert(varchar(32),'" + CStr(html2db(ArroptionName(i))) + "'), 'Y','Y', '" & limityn & "'," & CStr(optlimitno) & ", 0," & CStr(Trim(Arroptaddprice(i))) & "," & CStr(Trim(Arroptaddbuyprice(i))) & ") "
							'Response.write sqlStr
							'Response.end
							dbACADEMYget.Execute sqlStr
					End If
				Next
			ElseIf (Request.Form("optlevel")="2") then
				'' 이중옵션
				iErrMsg = WaitRegDoubleOptionProc(waititemid)
				if (iErrMsg<>"") then
					response.write iErrMsg
				end if
			end if

			''옵션 총수 저장
			sqlStr = "update db_academy.dbo.tbl_diy_wait_item"
			sqlStr = sqlStr + " set optioncnt=(select count(itemid) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y')"
			sqlStr = sqlStr + " , limitno=(select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y')"
			sqlStr = sqlStr + " ,limityn=(case"
			sqlStr = sqlStr + " when 0 < (select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y') then 'Y'"
			sqlStr = sqlStr + " when 0 >= (select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y') then 'N'"
			sqlStr = sqlStr + " end)"
			sqlStr = sqlStr + "where itemid = " + CStr(waititemid) + " "
			dbACADEMYget.Execute sqlStr
		End If
	ElseIf mode = "editOptionDel" Then
		If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
			Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
			Response.end
		End If
		'################### 기존 등록 옵션 삭제 #######################################################
		sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option"
		sqlStr = sqlStr & " where itemid=" & waititemid
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
	Else
		If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
			Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
			Response.end
		End If
		'###########################################################################
		'상품 데이터 수정
		'###########################################################################
		sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + vbCrlf
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
		sqlStr = sqlStr & " where itemid=" + CStr(waititemid) + vbCrlf
		dbACADEMYget.Execute sqlStr
		'###########################################################################
	End If
Else
	If mode = "editOption" Or mode = "editOptionMultiple" Then
		DesignerID = Request.Form("designerid")
		'###########################################################################
		'상품 데이터 입력
		'###########################################################################
		sqlStr = "insert into db_academy.dbo.tbl_diy_wait_item" + vbCrlf
		sqlStr = sqlStr & " (itemdiv,makerid,itemname,regdate,buycash, sellcash, mileage, sellyn, deliverytype,limityn,currstate)" + vbCrlf
		sqlStr = sqlStr & " values(" + vbCrlf
		sqlStr = sqlStr & "'01'" + vbCrlf
		sqlStr = sqlStr & ",'" & DesignerID & "'" + vbCrlf
		sqlStr = sqlStr & ",'tempitem'" + vbCrlf
		sqlStr = sqlStr & ",getdate()" + vbCrlf
		sqlStr = sqlStr & ",0" + vbCrlf
		sqlStr = sqlStr & ",0" + vbCrlf
		sqlStr = sqlStr & ",0" + vbCrlf
		sqlStr = sqlStr & ",'N'" + vbCrlf
		sqlStr = sqlStr & ",'9'" + vbCrlf
		sqlStr = sqlStr & ",'N'" + vbCrlf
		sqlStr = sqlStr & ",3)" + vbCrlf
		'Response.write sqlStr
		'Response.end
		dbACADEMYget.Execute sqlStr
		'###########################################################################
		'상품 아이디 가져오기
		'###########################################################################
		sqlStr = "Select IDENT_CURRENT('db_academy.dbo.tbl_diy_wait_item') as maxitemid "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			waititemid = rsACADEMYget("maxitemid")
		rsACADEMYget.close
		
		'###########################################################################
		'상품 옵션 넣기
		'###########################################################################
'Response.end
		redim ArroptionName(Request.Form("optionName").Count)
		redim Arroptaddprice(Request.Form("optaddprice").Count)
		redim Arroptaddbuyprice(Request.Form("optaddbuyprice").Count)
		
		optlimitno=0
		Arroptioncode = Split(Request.Form("optioncode"),",")
		ArrCnt = Request.Form("optionName").Count

		limityn = "N"
		if Request.Form("useoptionyn")="Y" then
			''단일 옵션
			if (Request.Form("optlevel")="1") Then
				If ArrCnt > 0 Then
					'################### 기존 등록 옵션 삭제 #######################################################
					sqlStr = "delete from db_academy.dbo.tbl_diy_wait_item_option"
					sqlStr = sqlStr & " where itemid=" & waititemid
					rsACADEMYget.Open sqlStr,dbACADEMYget,1
				End If

				for i=1 to ArrCnt
					ArroptionName(i) = Request.Form("optionName")(i)
					Arroptaddprice(i) = Request.Form("optaddprice")(i)
					Arroptaddbuyprice(i) = Request.Form("optaddbuyprice")(i)
					If optlimitno > 0 Then
						limityn = "Y"
					End If
					if (Trim(Arroptioncode(i-1)) <> "") then
							sqlStr = " insert into db_academy.dbo.tbl_diy_wait_item_option(itemid, itemoption, optionTypeName, optionname, isusing, optsellyn, optlimityn, optlimitno, optlimitsold , optaddPrice, optaddBuyPrice)"
							sqlStr = sqlStr + " values(" + CStr(waititemid) + ", '" + CStr(Trim(Arroptioncode(i-1))) + "', convert(varchar(32),'" + html2db(Request.Form("optionTypename")) + "'), convert(varchar(96),'" + CStr(html2db(ArroptionName(i))) + "'), 'Y','Y', '" & limityn & "'," & CStr(optlimitno) & ", 0," & CStr(Trim(Arroptaddprice(i))) & "," & CStr(Trim(Arroptaddbuyprice(i))) & ")"
							'Response.write sqlStr
							'Response.end
							dbACADEMYget.Execute sqlStr
					End If
				Next
			ElseIf (Request.Form("optlevel")="2") then
				'' 이중옵션
				iErrMsg = WaitRegDoubleOptionProc(waititemid)
				if (iErrMsg<>"") then
					response.write iErrMsg
				end if
			end if

			''옵션 총수 저장
			sqlStr = "update db_academy.dbo.tbl_diy_wait_item"
			sqlStr = sqlStr + " set optioncnt=(select count(itemid) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y')"
			sqlStr = sqlStr + " , limitno=(select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y')"
			sqlStr = sqlStr + " ,limityn=(case"
			sqlStr = sqlStr + " when 0 < (select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y') then 'Y'"
			sqlStr = sqlStr + " when 0 >= (select sum(optlimitno) as cnt from db_academy.dbo.tbl_diy_wait_item_option where itemid =" + CStr(waititemid) + " and isusing = 'Y') then 'N'"
			sqlStr = sqlStr + " end)"
			sqlStr = sqlStr + "where itemid = " + CStr(waititemid) + " "
			dbACADEMYget.Execute sqlStr
			mode = "editOptionFirst"
		End If
	End If
End If
%>
<script>
<!--
<% If mode = "editOptionDel" Then %>
	parent.fnDetailInfoEnd2();
<% elseIf mode = "editOptionFirst" Then %>
	parent.fnDetailInfoEnd3("<%=waititemid%>");
<% elseIf mode = "iteminfofirst" Then %>
	parent.fnDetailInfoEnd("<%=waititemid%>");
<% Else %>
	parent.fnDetailInfoEnd();
<% End If %>
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->