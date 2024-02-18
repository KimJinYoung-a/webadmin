<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%
dim itemid, largeno, midno, smallno, itemdiv
dim keywords, sourcearea, makername, itemsource
dim itemsize, itemWeight, usinghtml, itemcontent, ordercomment, designercomment, itemvideo
dim mode, itemoption, isusing, upchemanagecode
Dim cstodr , requireMakeDay , requirecontents , refundpolicy , infoDiv , safetyYn , safetyDiv , safetyNum
Dim freight_min , freight_max
Dim itemname
Dim requirechk , requireemail

itemid			= requestCheckvar(request("itemid"),10)
itemname        = html2db(requestCheckvar(request("itemname"),64))
largeno			= requestCheckvar(request("cd1"),10)
midno			= requestCheckvar(request("cd2"),10)
smallno			= requestCheckvar(request("cd3"),10)
itemdiv			= requestCheckvar(request("itemdiv"),2)
keywords		= html2db(request("keywords"))
sourcearea		= html2db(request("sourcearea"))
makername		= html2db(request("makername"))
itemsource		= html2db(request("itemsource"))
itemsize		= html2db(request("itemsize"))
itemWeight		= html2db(requestCheckvar(request("itemWeight"),10))
usinghtml		= requestCheckvar(request("usinghtml"),1)
itemcontent		= html2db(request("itemcontent"))
ordercomment	= html2db(request("ordercomment"))
designercomment	= html2db(request("designercomment"))
upchemanagecode = html2db(request("upchemanagecode"))

cstodr			= requestCheckvar(Request("cstodr"),1)
requireMakeDay	= requestCheckvar(Request("requireMakeDay"),10)
requirecontents	= html2db(Request("requirecontents"))
refundpolicy	= html2db(Request("refundpolicy"))
infoDiv			= requestCheckvar(Request("infoDiv"),2)
safetyYn		= requestCheckvar(Request("safetyYn"),1)
safetyDiv		= requestCheckvar(Request("safetyDiv"),10)
safetyNum		= chrbyte(html2db(Request("safetyNum")),24,"")
freight_min		= getNumeric(requestCheckvar(Request("freight_min"),10))
freight_max		= getNumeric(requestCheckvar(Request("freight_max"),10))

requirechk		= requestCheckvar(Request("requireimgchk"),1)
requireemail	= html2db(Request("requireMakeEmail"))

itemvideo       = Request("itemvideo")

if keywords <> "" then
	if checkNotValidHTML(keywords) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end if
if sourcearea <> "" then
	if checkNotValidHTML(sourcearea) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if makername <> "" then
	if checkNotValidHTML(makername) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if itemsource <> "" then
	if checkNotValidHTML(itemsource) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if itemsize <> "" then
	if checkNotValidHTML(itemsize) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if itemcontent <> "" then
	if checkNotValidHTML(itemcontent) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if ordercomment <> "" then
	if checkNotValidHTML(ordercomment) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if designercomment <> "" then
	if checkNotValidHTML(designercomment) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if upchemanagecode <> "" then
	if checkNotValidHTML(upchemanagecode) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if requirecontents <> "" then
	if checkNotValidHTML(requirecontents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if refundpolicy <> "" then
	if checkNotValidHTML(refundpolicy) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if requireemail <> "" then
	if checkNotValidHTML(requireemail) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end if
''권한체크
Dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid
oitem.FRectMakerId = session("ssBctID")
if (oitem.FRectMakerId<>"") then
    oitem.GetOneItem
end if
if (oitem.FResultCount < 1) then
    response.write "<script>alert('존재하지 않는 상품이거나 권한이 없습니다.'); self.close();</script>"
    dbACADEMYget.close()	:	response.End
end if


dim sqlStr,i
dim AssignedRow

'==============================================================================
sqlStr = "update db_academy.dbo.tbl_diy_item" + VbCrlf
sqlStr = sqlStr + " set cate_large='" + largeno + "'" + VbCrlf
sqlStr = sqlStr + " , cate_mid='" + midno + "'" + VbCrlf
sqlStr = sqlStr + " , cate_small='" + smallno + "'" + VbCrlf
sqlStr = sqlStr + " , itemdiv='" + CStr(itemdiv) + "'" + VbCrlf
sqlStr = sqlStr & " , upchemanagecode='" & LeftB(upchemanagecode,32) & "'" + vbCrlf
sqlStr = sqlStr + " , itemname=convert(varchar(64),'" + CStr(itemname) + "')" + VbCrlf
sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
sqlStr = sqlStr + " , requireimgchk = '" + requirechk + "'" + VbCrlf
sqlStr = sqlStr + " where itemid=" + CStr(itemid) + " "+ vbCrlf
sqlStr = sqlStr + " and makerid='" + CStr(session("ssBctID")) + "' "

dbACADEMYget.Execute sqlStr, AssignedRow

'// 상품상세 설명 유무에 따른 처리
	sqlStr = "select count(*) from db_academy.dbo.tbl_diy_item_Contents where itemid=" & itemid
	rsACADEMYget.Open sqlStr ,dbACADEMYget,1
	if rsACADEMYget(0)<1 then
	    sqlStr = "insert into db_academy.dbo.tbl_diy_item_Contents (itemid, keywords, sourcearea, makername, itemsource, itemsize, itemWeight, usinghtml, itemcontent, ordercomment, designercomment , cstodr , requireMakeDay , requirecontents,refundpolicy,infoDiv,safetyYn,safetyDiv,safetyNum,freight_min,freight_max , requiremakeemail) values " + VbCrlf
	    sqlStr = sqlStr + " (" + CStr(itemid) + VbCrlf
	    sqlStr = sqlStr + " , '" + keywords + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + sourcearea + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + makername + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + itemsource + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + itemsize + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + itemWeight + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + usinghtml + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + itemcontent + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + ordercomment + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + designercomment + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + cstodr + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + requireMakeDay + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + requirecontents + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + refundpolicy + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + infoDiv + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + safetyYn + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + safetyDiv + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + safetyNum + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + freight_min + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + freight_max + "'" + VbCrlf
	    sqlStr = sqlStr + " , '" + requireemail + "')" + VbCrlf
	    dbACADEMYget.Execute sqlStr
	else
	    sqlStr = "update db_academy.dbo.tbl_diy_item_Contents" + VbCrlf
	    sqlStr = sqlStr + " set keywords='" + keywords + "'" + VbCrlf
	    sqlStr = sqlStr + " , sourcearea='" + sourcearea + "'" + VbCrlf
	    sqlStr = sqlStr + " , makername='" + makername + "'" + VbCrlf
	    sqlStr = sqlStr + " , itemsource='" + itemsource + "'" + VbCrlf
	    sqlStr = sqlStr + " , itemsize='" + itemsize + "'" + VbCrlf
	    sqlStr = sqlStr + " , itemWeight='" + itemWeight + "'" + VbCrlf
	    sqlStr = sqlStr + " , usinghtml='" + usinghtml + "'" + VbCrlf
	    sqlStr = sqlStr + " , itemcontent='" + itemcontent + "'" + VbCrlf
	    sqlStr = sqlStr + " , ordercomment='" + ordercomment + "'" + VbCrlf
	    sqlStr = sqlStr + " , designercomment='" + designercomment + "'" + VbCrlf

		sqlStr = sqlStr & " ,cstodr='" & cstodr & "'" + vbCrlf
		sqlStr = sqlStr & " ,requireMakeDay='" & requireMakeDay & "'" + vbCrlf
		sqlStr = sqlStr & " ,requirecontents='" & requirecontents & "'" + vbCrlf
		sqlStr = sqlStr & " ,refundpolicy='" & refundpolicy & "'" + vbCrlf
		sqlStr = sqlStr & " ,infoDiv='" & infoDiv & "'" + vbCrlf
		sqlStr = sqlStr & " ,safetyYn='" & safetyYn & "'" + vbCrlf
		sqlStr = sqlStr & " ,safetyDiv='" & safetyDiv & "'" + vbCrlf
		sqlStr = sqlStr & " ,safetyNum='" & safetyNum & "'" + vbCrlf
		sqlStr = sqlStr & " ,freight_min='" & freight_min & "'" + vbCrlf
		sqlStr = sqlStr & " ,freight_max='" & freight_max & "'" + vbCrlf
		sqlStr = sqlStr & " ,requiremakeemail='" & requireemail & "'" + vbCrlf

	    sqlStr = sqlStr + " where itemid=" + CStr(itemid) + " "
	    dbACADEMYget.Execute sqlStr
	end if
	rsACADEMYget.Close

	'###########################################################################
	'상품 품목고시정보 저장 
	'###########################################################################
	if Request("infoDiv")<>"" then
		dim infoCd, infoCont, infoChk
	
		'배열로 처리
		redim infoCd(Request("infoCd").Count)
		redim infoCont(Request("infoCont").Count)
		redim infoChk(Request("infoChk").Count)
		for i=1 to Request("infoCd").Count
			infoCd(i) = Request("infoCd")(i)
			infoCont(i) = Request("infoCont")(i)
			infoChk(i) = Request("infoChk")(i)
		next
	
		'기존값 삭제
		sqlStr = "Delete From db_academy.dbo.tbl_diy_item_infoCont Where itemid='" & CStr(itemid) & "'"&VbCRLF
		dbACADEMYget.execute(sqlStr)
	
		'DB에 처리
		for i=1 to ubound(infoCd)
			'입력값이 있는 경우만 저장
			if infoChk(i)<>"" or infoCont(i)<>"" then
				sqlStr = "Insert into db_academy.dbo.tbl_diy_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
				sqlStr = sqlStr & "('" & CStr(itemid) & "'"
				sqlStr = sqlStr & ",'" & CStr(infoCd(i)) & "'"
				sqlStr = sqlStr & ",'" & CStr(infoChk(i)) & "'"
				sqlStr = sqlStr & ",'" & html2db(infoCont(i)) & "')"
				dbACADEMYget.execute(sqlStr)
			end if
		Next
	end If


	'// 전시카테고리 넣기 //
	sqlStr = "delete from db_academy.dbo.tbl_display_cate_item_Academy Where itemid='" & itemid & "';" & vbCrLf
	If (Request("catecode").Count>0) Then
		sqlStr = sqlStr & "update db_academy.dbo.tbl_diy_item set dispcate1=null Where itemid='" & itemid & "';" & vbCrLf
		for i=1 to Request("catecode").Count
			sqlStr = sqlStr & "Insert into db_academy.dbo.tbl_display_cate_item_Academy (catecode, itemid, depth, sortNo, isDefault) values "
			sqlStr = sqlStr & "('" & Request("catecode")(i) & "'"
			sqlStr = sqlStr & ",'" & itemid & "'"
			sqlStr = sqlStr & ",'" & Request("catedepth")(i) & "',9999"
			sqlStr = sqlStr & ",'" & Request("isDefault")(i) & "');" & vbCrLf
			
			if Request("isDefault")(i)="y" then
				sqlStr = sqlStr & "update db_academy.dbo.tbl_diy_item set dispcate1='" & left(Request("catecode")(i),3) & "' Where itemid='" & itemid & "';" & vbCrLf
			end if
		next
	end if
	dbACADEMYget.execute(sqlStr)
	
if (AssignedRow>0) then
	'// 카테고리 중복 확인(2008.07.31; 허진원)
	sqlStr = "select count(*) from db_academy.dbo.tbl_diy_item_category where itemid=" & itemid &VbCRLF
	sqlStr = sqlStr & "	and code_large='" & largeno & "' " &VbCRLF
	sqlStr = sqlStr & "	and code_mid='" & midno & "' " &VbCRLF
	sqlStr = sqlStr & "	and code_small='" & smallno & "' and code_div='A' "
	rsACADEMYget.Open sqlStr ,dbACADEMYget,1

	if rsACADEMYget(0)<1 then
	    '''新 카테고리 : 업체는 기본 카테고리만 가능
	    sqlStr = "update db_academy.dbo.tbl_diy_item_category " 
	    sqlStr = sqlStr + " set code_large='" + largeno + "'"
	    sqlStr = sqlStr + " , code_mid='" + midno + "'"
	    sqlStr = sqlStr + " , code_small='" + smallno + "'"
	    sqlStr = sqlStr + " where itemid=" & CStr(itemid)
	    sqlStr = sqlStr + " and code_div='D'"
	    sqlStr = sqlStr + " and ("
	    sqlStr = sqlStr + "         code_large<>'" + largeno + "'"
	    sqlStr = sqlStr + "     or  code_mid<>'" + midno + "'"
	    sqlStr = sqlStr + "     or  code_small<>'" + smallno + "'"
	    sqlStr = sqlStr + " )"

	    dbACADEMYget.Execute sqlStr
	else
		Response.Write "<script language=javascript>alert('이미 상품에 지정되어있는 카테고리를 선택하였습니다.\n\n※추가 카테고리가 지정되어있을 경우가 있으므로 담당MD에게 확인/수정요청을 해주세요.');history.back();</script>"
		dbACADEMYget.close()	:	response.End
	end if

	rsACADEMYget.Close
end if

    If Trim(itemvideo) <> "" Then
		Dim RetStr, RetSrc, RetWidth, RetHeight, regEx, Matches, Match, VideoTempSrc, VideoTempWidth, VideoTempHeight, videoType, dbsql
		itemvideo = Trim(Replace(itemvideo,"""","'"))
		itemvideo = replace(itemvideo,"BUFiframe","iframe")
		
		'// iframe 이외의 코드는 잘라버림
		itemvideo = Left(itemvideo, InStrRev(itemvideo, "</iframe>")+9)

		'// 비디오 타입지정(유투브인지 비메오인지)
		If InStr(itemvideo, "youtube")>0 Then
			videoType = "youtube"
		ElseIf InStr(itemvideo, "vimeo")>0 Then
			videoType = "vimeo"
		Else
			videoType = "etc"
		End If

		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True

		regEx.pattern = "<iframe [^<>]*>"
		Set Matches = regEx.execute(itemvideo)
		For Each Match In Matches
			VideoTempSrc =  Mid(Match.Value, InStrRev(Match.Value,"src='")+5)
			RetSrc = Left(VideoTempSrc, InStr(VideoTempSrc, "'")-1)
			
			If InStrRev(Match.Value,"width='") > 0 then
			VideoTempWidth =  Mid(Match.Value, InStrRev(Match.Value,"width='")+7)
			RetWidth = Left(VideoTempWidth, InStr(VideoTempWidth, "'")-1)
			End If 
			
			If InStrRev(Match.Value,"height='") > 0 then
			VideoTempHeight =  Mid(Match.Value, InStrRev(Match.Value,"height='")+8)
			RetHeight = Left(VideoTempHeight, InStr(VideoTempHeight, "'")-1)
			End If 
		Next
		Set regEx = Nothing
		Set Matches = Nothing

		sqlStr = " select idx FROM [db_academy].[dbo].[tbl_diy_item_videos]  WHERE videogubun='video1' And itemid =" + CStr(itemid)
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if Not rsACADEMYget.Eof Then
			If Not(videoType="etc") Then
				'// 데이터가 있다면 업데이트 해줌.
				dbsql = "update [db_academy].[dbo].[tbl_diy_item_videos] " + vbCrlf
				dbsql = dbsql & " set videourl='" &RetSrc& "'" + vbCrlf
				dbsql = dbsql & " ,videowidth='" & RetWidth & "'" + vbCrlf
				dbsql = dbsql & " ,videoheight='" & RetHeight & "'" + vbCrlf
				dbsql = dbsql & " ,videotype='" & videoType & "'" + vbCrlf
				dbsql = dbsql & " ,videofullurl='" & chrbyte(html2db(itemvideo),255,"") & "'" + vbCrlf
				dbsql = dbsql & " ,modifydate=getdate()" + vbCrlf
				dbsql = dbsql & " where idx='"&rsACADEMYget("idx")&"' And itemid='" & CStr(itemid) & "'" + vbCrlf
				dbACADEMYget.execute(dbsql)
			End If
		Else
			If Not(videoType="etc") Then
				'// 데이터가 없으면 인서트 해줌.
				dbsql = " insert into [db_academy].[dbo].[tbl_diy_item_videos]  (itemid, videogubun, videotype, videourl, videowidth, videoheight, videofullurl, regdate) values " + vbCrlf
				dbsql = dbsql & " ('"&CStr(itemid)&"', 'video1', '"&videoType&"', '"&RetSrc&"', '"&RetWidth&"', '"&RetHeight&"','"&chrbyte(html2db(itemvideo),255,"")&"', getdate()) " + vbCrlf
				dbACADEMYget.execute(dbsql)
			End If
		end if
		rsACADEMYget.close
	Else
		'// 아무값도 안넘어왔는데 db에 값이 있으면 삭제라고 판단. 지워줌.
		sqlStr = " select idx FROM [db_academy].[dbo].[tbl_diy_item_videos]  WHERE videogubun='video1' And itemid =" + CStr(itemid)  
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if Not rsACADEMYget.Eof Then
			dbsql = " Delete from [db_academy].[dbo].[tbl_diy_item_videos]  Where videogubun='video1' And itemid=" + CStr(itemid)
			dbACADEMYget.execute(dbsql)
		End If
		rsACADEMYget.close
	End If

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->