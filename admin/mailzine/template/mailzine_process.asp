<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbTMSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinenewcls.asp"-->
<!-- #include virtual="/lib/classes/search/searchMobileCls.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/mailzine/template/mailzine_html_make.asp"-->
<%
dim idx , imgnumber , mode , sql, POST_ID, mailcontents1, mailcontents2, yyyymmdd, mailergubun
dim regtype, regdate, evt_code, evtList, itemList, itemList2, itemList3, sqlStr, affectedRows, member
dim title,img1editname,img2editname,img3editname,img4editname,area,gubun,memgubun,secretgubun, just1day, tentenclass
dim classDate, itemid1, salePer1, classDesc1, classSubDesc1, itemid2, salePer2, classDesc2, classSubDesc2, itemid3
dim cMailzine, ArrTemplateInfo, ix, codesplit, masteridx, salePer3, classDesc3, classSubDesc3
dim contentsTemp
	mailergubun = requestcheckvar(request("mailergubun"),16)
	idx 		= requestCheckVar(request("idx"),32)
	imgnumber 	= requestCheckVar(request("imgnumber"),32)
	mode 		= requestCheckVar(request("mode"), 32)
	member 		= requestCheckVar(request("member"), 64)
	yyyymmdd 	= requestCheckVar(request("yyyymmdd"), 64)

	regtype = requestCheckVar(request("regtype"), 32)
	regdate = requestCheckVar(request("regdate"), 32)
	evt_code = requestCheckVar(request("evt_code"), 300)
	arrevtcode = requestCheckVar(request("arrevtcode"), 300)

	title 			= requestCheckVar(request("title"), 3200)
	img1editname 	= requestCheckVar(request("img1editname"), 3200)
	img2editname 	= requestCheckVar(request("img2editname"), 3200)
	img3editname 	= requestCheckVar(request("img3editname"), 3200)
	img4editname 	= requestCheckVar(request("img4editname"), 3200)
	area 			= requestCheckVar(request("area"), 3200)
	gubun 			= requestCheckVar(request("gubun"), 3200)
	memgubun 		= requestCheckVar(request("memgubun"), 3200)
	secretgubun 	= requestCheckVar(request("secretgubun"), 3200)

	classDate 	= requestCheckVar(request("classDate"), 10)
	itemid1 	= requestCheckVar(request("itemid1"), 10)
	salePer1 	= requestCheckVar(request("salePer1"), 10)
	classDesc1 	= requestCheckVar(request("classDesc1"), 64)
	classSubDesc1 	= requestCheckVar(request("classSubDesc1"), 64)
	itemid2 	= requestCheckVar(request("itemid2"), 10)
	salePer2 	= requestCheckVar(request("salePer2"), 10)
	classDesc2 	= requestCheckVar(request("classDesc2"), 64)
	classSubDesc2 	= requestCheckVar(request("classSubDesc2"), 64)
	itemid3 	= requestCheckVar(request("itemid3"), 10)
	salePer3 	= requestCheckVar(request("salePer3"), 10)
	classDesc3 	= requestCheckVar(request("classDesc3"), 64)
	classSubDesc3 	= requestCheckVar(request("classSubDesc3"), 64)

	if evt_code <> "" then
		if instr(evt_code,",") then
			codesplit = split(evt_code,",")
			evt_code = codesplit(0)
		else
			evt_code=evt_code
		end if	
	else
		if instr(arrevtcode,",") then
			codesplit = split(arrevtcode,",")
			evt_code = codesplit(0)
		else
			evt_code=arrevtcode
		end if
	end if
if mode = "imgdel" then

	if idx = "" or imgnumber = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('idx 값이나 삭제할 이미지 번호가 없습니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.end
	end if

	sql = "update [db_sitemaster].[dbo].tbl_mailzine set" + vbcrlf

	if imgnumber = "1" then
		sql = sql & " img1 = null" + vbcrlf
		sql = sql & " ,imgmap1 = null" + vbcrlf
	elseif imgnumber = "2" then
		sql = sql & " img2 = null" + vbcrlf
		sql = sql & " ,imgmap2 = null" + vbcrlf
	elseif imgnumber = "3" then
		sql = sql & " img3 = null" + vbcrlf
		sql = sql & " ,imgmap3 = null" + vbcrlf
	elseif imgnumber = "4" then
		sql = sql & " img4 = null" + vbcrlf
		sql = sql & " ,imgmap4 = null" + vbcrlf
	end if

	sql = sql & " where idx = "&idx&""

	'response.write sql &"<Br>"
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	location.href='/admin/mailzine/template/mailzine_detail.asp?idx="&idx&"';"
	response.write "</script>"
	dbget.close()	:	response.end

'// 가져오기
elseif (mode = "getlist") then

	''템플릿 정보 가져오기
	dim arrEvtCode, scriptTXT, arrJust1day, arrMdpick, debugTxt, debugTxt2
	set cMailzine = new CMailzineList
	cMailzine.FRectRegType = regtype
	ArrTemplateInfo=cMailzine.fnMailzineTemplateInfo
	set cMailzine = nothing

	If isArray(ArrTemplateInfo) Then
		if evt_code = "" then evt_code="0"
		arrEvtCode=evt_code
		arrJust1day = 0
		arrMdpick = 0
		scriptTXT=""
		debugTxt=""
		debugTxt2=""
		For ix=0 To UBound(ArrTemplateInfo,2)
			debugTxt2 = debugTxt2 & "//" & ArrTemplateInfo(0,ix) & " / " & Cstr(ix) & vbcrlf
			if ArrTemplateInfo(0,ix)="20" or ArrTemplateInfo(0,ix)="21" or ArrTemplateInfo(0,ix)="22" or ArrTemplateInfo(0,ix)="23" or ArrTemplateInfo(0,ix)="24" or ArrTemplateInfo(0,ix)="25" then
				'수작업 영역, 메인 이벤트코드 제외
			else
				if ArrTemplateInfo(0,ix)="26" then '이벤트 목록
					evtList=""
					evtList = GetEventFromMainBanner(regdate, arrEvtCode, ArrTemplateInfo(2,ix))
					if arrEvtCode <> "" then
						if evtList<>"" then
							arrEvtCode = arrEvtCode + "," + evtList
						end if
					else
						arrEvtCode = evtList
					end if
					scriptTXT = scriptTXT + "	parent.document.frm.contents"&ix+1&".value = '" & Replace(evtList, ",", "\n") & "';" & vbcrlf
					debugTxt = debugTxt & "//evtlist-" & ArrTemplateInfo(0,ix) & " [" & evtList & "] / " & Cstr(ix) & vbcrlf
				elseif ArrTemplateInfo(0,ix)="27" then 'MDPick
					itemList=""
					itemList = GetMDPick(regdate, arrMdpick, ArrTemplateInfo(2,ix))
					if arrMdpick <> "" then
						if itemList<>"" then
							arrMdpick = Cstr(arrMdpick) + "," + Cstr(itemList)
						end if
					else
						arrMdpick = itemList
					end if
					scriptTXT = scriptTXT + "	parent.document.frm.contents"&ix+1&".value = '" & Replace(itemList, ",", "\n") & "';" & vbcrlf
					debugTxt = debugTxt & "//mdpick-" & ArrTemplateInfo(0,ix) & " [" & itemList & "] / " & Cstr(ix) & vbcrlf
				elseif ArrTemplateInfo(0,ix)="28" then 'New
					itemList2=""
					itemList2 = GetMDPickNew(regdate, ArrTemplateInfo(2,ix))
					scriptTXT = scriptTXT + "	parent.document.frm.contents"&ix+1&".value = '" & Replace(itemList2, ",", "\n") & "';" & vbcrlf
					debugTxt = debugTxt & "//new-" & ArrTemplateInfo(0,ix) & " [" & itemList2 & "] / " & Cstr(ix) & vbcrlf
				elseif ArrTemplateInfo(0,ix)="29" then 'Best
					itemList3=""
					itemList3 = GetMDPickBest(regdate, ArrTemplateInfo(2,ix))
					scriptTXT = scriptTXT + "	parent.document.frm.contents"&ix+1&".value = '" & Replace(itemList3, ",", "\n") & "';" & vbcrlf
					debugTxt = debugTxt & "//best-" & ArrTemplateInfo(0,ix) & " [" & itemList3 & "] / " & Cstr(ix) & vbcrlf
				elseif ArrTemplateInfo(0,ix)="30" then 'just1day
					just1day=""
					just1day = GetJustOneDay2018(regdate, arrJust1day, ArrTemplateInfo(2,ix))
					if arrJust1day <> "" then
						if just1day<>"" then
							arrJust1day = Cstr(arrJust1day) + "," + Cstr(just1day)
						end if
					else
						arrJust1day = just1day
					end if
					scriptTXT = scriptTXT + "	parent.document.frm.contents"&ix+1&".value = '" & Replace(just1day, ",", "\n") & "';" & vbcrlf
					debugTxt = debugTxt & "//just1day-" & ArrTemplateInfo(0,ix) & " [" & just1day & "] / " & Cstr(ix) & vbcrlf
				elseif ArrTemplateInfo(0,ix)="31" then 'Class
					tentenclass=""
					tentenclass = GetTenTenClass(regdate)
					scriptTXT = scriptTXT + "	parent.document.frm.contents"&ix+1&".value = '" & Replace(tentenclass, ",", "\n") & "';" & vbcrlf
					debugTxt = debugTxt & "//class-" & ArrTemplateInfo(0,ix) & " [" & tentenclass & "] / " & Cstr(ix) & vbcrlf
				end if
			end if
			
		Next
	end if

	response.write "<script type='text/javascript'>" & vbcrlf
	response.write scriptTXT
	response.write "	alert('OK');" & vbcrlf
	response.write debugTxt
	response.write debugTxt2
	response.write "</script>"

' 신규등록
elseif (mode = "ins") then
	if mailergubun="" or isnull(mailergubun) then
		response.write "메일러 구분이 없습니다."
		dbget.close() : response.end
	end if

	if (evt_code = "") then
		evt_code = "NULL"
	end if
	dim iA ,arrTemp,arrItemid, itemid
	'템플릿 정보 가져오기
	set cMailzine = new CMailzineList
	cMailzine.FRectRegType = regtype
	ArrTemplateInfo=cMailzine.fnMailzineTemplateInfo
	set cMailzine = nothing

	sqlStr = "insert into [db_sitemaster].[dbo].tbl_mailzine" + vbcrlf
	sqlStr = sqlStr & " (title,regdate,area,gubun,memgubun,secretgubun,insertDate, reguserid, regtype2, evt_code, mailergubun)" + vbcrlf
	sqlStr = sqlStr & " values(" + vbcrlf
	sqlStr = sqlStr & "'" & html2db(title) & "'," + vbcrlf
	sqlStr = sqlStr & "'" & regdate & "'," + vbcrlf
	sqlStr = sqlStr & "'" & area & "'," + vbcrlf
	sqlStr = sqlStr & "" & gubun & "," + vbcrlf
	sqlStr = sqlStr & "'" & memgubun & "'," + vbcrlf        ''추가
	sqlStr = sqlStr & "'" & secretgubun & "'," + vbcrlf        ''2013-10-01 추가
	sqlStr = sqlStr & " getdate(), " + vbcrlf        '2013-12-27 김진영 추가
	sqlStr = sqlStr & " '" & session("ssBctId") & "', " + vbcrlf
	sqlStr = sqlStr & " '" & regtype & "', " + vbcrlf
	sqlStr = sqlStr & " " & evt_code & " " + vbcrlf
	sqlStr = sqlStr & " ,'" & mailergubun & "'" & vbcrlf
	sqlStr = sqlStr & ")"

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	sqlStr = "select SCOPE_IDENTITY()"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	masteridx = rsget(0)
	rsget.Close

	If isArray(ArrTemplateInfo) Then
		For ix=0 To UBound(ArrTemplateInfo,2)

		' 이미지배너01 ~ 이미지배너04
		if ArrTemplateInfo(0, ix)="20" or ArrTemplateInfo(0, ix)="21" or ArrTemplateInfo(0, ix)="22" or ArrTemplateInfo(0, ix)="23" then 
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
			sqlStr = sqlStr & " (masteridx, contents, img, contentsCode, codeidx)" & vbcrlf
			sqlStr = sqlStr & " values(" & vbcrlf
			sqlStr = sqlStr & " " & masteridx & "" & vbcrlf
			sqlStr = sqlStr & " ,'" & html2db(request("imagemap"&ix+1)) & "'" & vbcrlf
			sqlStr = sqlStr & " ,'" & request("img"&ix+1) & "','" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(3, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " )"
		' 주말특가
		elseif ArrTemplateInfo(0,ix)="24" then
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
			sqlStr = sqlStr & " (masteridx, contents, img, contentsCode, codeidx)" & vbcrlf
			sqlStr = sqlStr & " values(" & vbcrlf
			sqlStr = sqlStr & " " & masteridx & "" & vbcrlf
			sqlStr = sqlStr & " ,'" & html2db(request("evt_code"&ix+1)) & "',''" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(3, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " )"
		' 메인이벤트코드
		elseif ArrTemplateInfo(0,ix)="25" then
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
			sqlStr = sqlStr & " (masteridx, contents, img, contentsCode, codeidx)" & vbcrlf
			sqlStr = sqlStr & " values(" & vbcrlf
			sqlStr = sqlStr & " " & masteridx & "" & vbcrlf
			sqlStr = sqlStr & " ,'" & html2db(request("evt_code"&ix+1)) & "',''" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(3, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " )"
		' MDPICK
		elseif ArrTemplateInfo(0,ix)="27" then
			arrTemp = replace(trim(request("contents"&ix+1)),chr(13),"")
			arrTemp = replace(arrTemp,chr(10),",")
			if right(arrTemp,1)="," then arrTemp = left(arrTemp,len(arrTemp)-1)

			sqlStr = "select"
			sqlStr = sqlStr & "	itemid"
			sqlStr = sqlStr & "	from db_item.dbo.tbl_item with (nolock)"
			sqlStr = sqlStr & "	where isnull(adultType,0)=0"		' 성인용품제외
			sqlStr = sqlStr & "	and itemid in ("& arrTemp &")"

			'response.write sqlStr & "<Br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

			if not rsget.EOF then
				do until rsget.eof
					contentsTemp = contentsTemp & rsget("itemid") & ","
					rsget.moveNext
				loop
			end if
			rsget.Close
			contentsTemp=trim(contentsTemp)
			if right(contentsTemp,1)="," then contentsTemp = left(contentsTemp,len(contentsTemp)-1)
			contentsTemp = replace(contentsTemp,",",vbcrlf)
			arrTemp = ""

			sqlStr = "insert into [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
			sqlStr = sqlStr & " (masteridx, contents, img, contentsCode, codeidx)" & vbcrlf
			sqlStr = sqlStr & " values(" & vbcrlf
			sqlStr = sqlStr & " " & masteridx & "" & vbcrlf
			sqlStr = sqlStr & " ,'" & html2db(contentsTemp) & "',''" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(3, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " )"
		else
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
			sqlStr = sqlStr & " (masteridx, contents, img, contentsCode, codeidx)" & vbcrlf
			sqlStr = sqlStr & " values(" & vbcrlf
			sqlStr = sqlStr & " " & masteridx & "" & vbcrlf
			sqlStr = sqlStr & " ,'" & html2db(request("contents"&ix+1)) & "',''" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " ,'" & ArrTemplateInfo(3, ix) & "'" & vbcrlf
			sqlStr = sqlStr & " )"
		end if

		'response.write sqlStr & "<br>"
		dbget.execute sqlStr

			'Best & New 상품 컨텐츠 등록
			if ArrTemplateInfo(0,ix)="28" or ArrTemplateInfo(0,ix)="29" then
				if ArrTemplateInfo(0,ix)="28" then gubun="N"
				if ArrTemplateInfo(0,ix)="29" then gubun="B"
				itemid = request("contents"&ix+1)
				itemid = replace(itemid,chr(13),"")
				arrTemp = Split(itemid,chr(10))
				sqlStr = "delete from [db_sitemaster].[dbo].tbl_main_mdchoice_Best_New where startdate>='" & Cstr(regdate & " 00:00:00") & "' and enddate<='" & Cstr(regdate & " 23:59:59") & "'" + vbcrlf
				dbget.execute sqlStr
				iA = 0
				do while iA <= ubound(arrTemp)
					if trim(arrTemp(iA))<>"" then
						sqlStr = "insert into [db_sitemaster].[dbo].tbl_main_mdchoice_Best_New("
						sqlStr = sqlStr & " linkinfo, linkitemid, disporder, startdate, enddate, gubun"
						sqlStr = sqlStr & " )"
						sqlStr = sqlStr & " 	select"
						sqlStr = sqlStr & " 	'" & "/shopping/category_prd.asp?itemid="& arrTemp(iA) & "'," & arrTemp(iA) & "," & Cstr(iA+1) & ""
						sqlStr = sqlStr & " 	,'" & Cstr(regdate & " 00:00:00") & "','" & Cstr(regdate & " 23:59:59") & "','" & gubun & "'"
						sqlStr = sqlStr & " 	from db_item.dbo.tbl_item with (nolock)"
						sqlStr = sqlStr & " 	where isnull(adultType,0)=0"		' 성인용품제외
						sqlStr = sqlStr & " 	and itemid="& arrTemp(iA) &""

						'response.write sqlStr & "<br>"
						dbget.execute sqlStr
					end if
					iA = iA + 1
				loop
			end if

		Next
	end if

	response.write "<script type='text/javascript'>"
	response.write "	alert('신규 저장 되었습니다.');"
	response.write "	opener.location.reload();"
	response.write "	opener.focus();"
	response.write "	window.close();"
	response.write "</script>"

' 수정
elseif (mode = "modi") then
	if mailergubun="" or isnull(mailergubun) then
		response.write "메일러 구분이 없습니다."
		dbget.close() : response.end
	end if

	if (evt_code = "") then
		evt_code = "NULL"
	end if

	sqlStr = " update [db_sitemaster].[dbo].tbl_mailzine "
	sqlStr = sqlStr & " set lastUpdate = getdate() "
	sqlStr = sqlStr & " , modiuserid = '" & session("ssBctId") & "' "
	sqlStr = sqlStr & " , title = '" & html2db(title) & "' "
	sqlStr = sqlStr & " , regdate = '" & regdate & "' "
	sqlStr = sqlStr & " , area = '" & area & "' "
	sqlStr = sqlStr & " , gubun = '" & gubun & "' "
	sqlStr = sqlStr & " , memgubun = '" & memgubun & "' "
	sqlStr = sqlStr & " , secretgubun = '" & secretgubun & "' "
	sqlStr = sqlStr & " , regtype2 = '" & regtype & "' "
	sqlStr = sqlStr & " , evt_code = " & evt_code & " "
	sqlStr = sqlStr & " , mailergubun = '" & mailergubun & "' where" & vbcrlf
	sqlStr = sqlStr & " idx = " & idx

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr, affectedRows

	'템플릿 정보 가져오기
	set cMailzine = new CMailzineList
	cMailzine.FRectRegType = regtype
	ArrTemplateInfo=cMailzine.fnMailzineTemplateInfo
	set cMailzine = nothing

	If isArray(ArrTemplateInfo) Then
		For ix=0 To UBound(ArrTemplateInfo,2)
			if request("idx"&ix+1)<>"" then
				' 이미지배너01 ~ 이미지배너04
				if ArrTemplateInfo(0, ix)="20" or ArrTemplateInfo(0, ix)="21" or ArrTemplateInfo(0, ix)="22" or ArrTemplateInfo(0, ix)="23" then 
					sqlStr = "update [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
					sqlStr = sqlStr & " set contents='" & html2db(request("imagemap"&ix+1)) & "'" & vbcrlf
					sqlStr = sqlStr & " , img='" & request("img"&ix+1) & "'" & vbcrlf
					sqlStr = sqlStr & " , contentsCode='" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
					sqlStr = sqlStr & " where masteridx=" & idx
					sqlStr = sqlStr & " and idx=" & request("idx"&ix+1)
				' 주말특가
				elseif ArrTemplateInfo(0,ix)="24" then
					sqlStr = "update [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
					sqlStr = sqlStr & " set contents='" & html2db(request("evt_code"&ix+1)) & "'" & vbcrlf
					sqlStr = sqlStr & " , contentsCode='" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
					sqlStr = sqlStr & " where masteridx=" & idx
					sqlStr = sqlStr & " and idx=" & request("idx"&ix+1)
				' 메인이벤트코드
				elseif ArrTemplateInfo(0,ix)="25" then
					sqlStr = "update [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
					sqlStr = sqlStr & " set contents='" & html2db(request("evt_code"&ix+1)) & "'" & vbcrlf
					sqlStr = sqlStr & " , contentsCode='" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
					sqlStr = sqlStr & " where masteridx=" & idx
					sqlStr = sqlStr & " and idx=" & request("idx"&ix+1)
				' MDPICK
				elseif ArrTemplateInfo(0,ix)="27" then
					arrTemp = replace(trim(request("contents"&ix+1)),chr(13),"")
					arrTemp = replace(arrTemp,chr(10),",")
					if right(arrTemp,1)="," then arrTemp = left(arrTemp,len(arrTemp)-1)

					sqlStr = "select"
					sqlStr = sqlStr & "	itemid"
					sqlStr = sqlStr & "	from db_item.dbo.tbl_item with (nolock)"
					sqlStr = sqlStr & "	where isnull(adultType,0)=0"		' 성인용품제외
					sqlStr = sqlStr & "	and itemid in ("& arrTemp &")"

					'response.write sqlStr & "<Br>"
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

					if not rsget.EOF then
						do until rsget.eof
							contentsTemp = contentsTemp & rsget("itemid") & ","
							rsget.moveNext
						loop
					end if
					rsget.Close
					contentsTemp=trim(contentsTemp)
					if right(contentsTemp,1)="," then contentsTemp = left(contentsTemp,len(contentsTemp)-1)
					contentsTemp = replace(contentsTemp,",",vbcrlf)
					arrTemp = ""

					sqlStr = "update [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
					sqlStr = sqlStr & " set contents='" & html2db(contentsTemp) & "'" & vbcrlf
					sqlStr = sqlStr & " , contentsCode='" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
					sqlStr = sqlStr & " where masteridx=" & idx
					sqlStr = sqlStr & " and idx=" & request("idx"&ix+1)
				else
					sqlStr = "update [db_sitemaster].[dbo].tbl_mailzine_contents" & vbcrlf
					sqlStr = sqlStr & " set contents='" & html2db(request("contents"&ix+1)) & "'" & vbcrlf
					sqlStr = sqlStr & " , contentsCode='" & ArrTemplateInfo(0, ix) & "'" & vbcrlf
					sqlStr = sqlStr & " where masteridx=" & idx
					sqlStr = sqlStr & " and idx=" & request("idx"&ix+1)
				end if

				'response.write sqlStr & "<br>"
				dbget.execute sqlStr

				'Best & New 상품 컨텐츠 등록
				if ArrTemplateInfo(0,ix)="28" or ArrTemplateInfo(0,ix)="29" then
					if ArrTemplateInfo(0,ix)="28" then gubun="N"
					if ArrTemplateInfo(0,ix)="29" then gubun="B"
					itemid = request("contents"&ix+1)
					itemid = replace(itemid,chr(13),"")
					arrTemp = Split(itemid,chr(10))
					sqlStr = "delete from [db_sitemaster].[dbo].tbl_main_mdchoice_Best_New where startdate>='" & Cstr(regdate & " 00:00:00") & "' and enddate<='" & Cstr(regdate & " 23:59:59") & "' and gubun='" & gubun & "'" + vbcrlf
					dbget.execute sqlStr
					iA = 0
					do while iA <= ubound(arrTemp)
						if trim(arrTemp(iA))<>"" then
							sqlStr = "insert into [db_sitemaster].[dbo].tbl_main_mdchoice_Best_New("
							sqlStr = sqlStr & " linkinfo, linkitemid, disporder, startdate, enddate, gubun"
							sqlStr = sqlStr & " )"
							sqlStr = sqlStr & " 	select"
							sqlStr = sqlStr & " 	'" & "/shopping/category_prd.asp?itemid="& arrTemp(iA) & "'," & arrTemp(iA) & "," & Cstr(iA+1) & ""
							sqlStr = sqlStr & " 	,'" & Cstr(regdate & " 00:00:00") & "','" & Cstr(regdate & " 23:59:59") & "','" & gubun & "'"
							sqlStr = sqlStr & " 	from db_item.dbo.tbl_item with (nolock)"
							sqlStr = sqlStr & " 	where isnull(adultType,0)=0"		' 성인용품제외
							sqlStr = sqlStr & " 	and itemid="& arrTemp(iA) &""

							'response.write sqlStr & "<br>"
							dbget.execute sqlStr
						end if
						iA = iA + 1
					loop
				end if
			end if
		Next
	end if
	response.write "<script type='text/javascript'>"
	if (affectedRows = 1) then
		response.write "	alert('수정 되었습니다.');"
	else
		response.write "	alert('----------------------------------------------------------\n\nERROR\n\n----------------------------------------------------------');"
	end if
	response.write "	opener.location.reload();"
	response.write "	opener.focus();"
	response.write "	window.close();"
	response.write "</script>"

elseif (mode = "inscls") then
	sqlStr = " insert into [db_sitemaster].[dbo].[tbl_mailzine_class]( "
	sqlStr = sqlStr & " 	classDate, itemid1, salePer1, classDesc1, classSubDesc1 "
	if (itemid2 <> "") and (itemid3 <> "") then
		sqlStr = sqlStr & " 	, itemid2, salePer2, classDesc2, classSubDesc2, itemid3, salePer3, classDesc3, classSubDesc3 "
	end if
	sqlStr = sqlStr & " 	, regdate, reguserid "
	sqlStr = sqlStr & " ) "
	sqlStr = sqlStr & " values('" & classDate & "', " & itemid1 & ", " & salePer1 & ", '" & classDesc1 & "', '" & classSubDesc1 & "' "
	if (itemid2 <> "") and (itemid3 <> "") then
		sqlStr = sqlStr & " , " & itemid2 & ", " & salePer2 & ", '" & classDesc2 & "', '" & classSubDesc2 & "', " & itemid3 & ", " & salePer3 & ", '" & classDesc3 & "', '" & classSubDesc3 & "' "
	end if
	sqlStr = sqlStr & " , getdate(), '" & session("ssBctId") & "') "

	''response.write sqlStr&"<br>"
	''response.end
	dbget.execute sqlStr, affectedRows

	response.write "<script type='text/javascript'>"
	if (affectedRows = 1) then
		response.write "	alert('OK');"
	else
		response.write "	alert('----------------------------------------------------------\n\nERROR\n\n----------------------------------------------------------');"
	end if
	response.write "	opener.location.reload();"
	response.write "	opener.focus();"
	response.write "	window.close();"
	response.write "</script>"

elseif (mode = "modicls") then

	sqlStr = " update [db_sitemaster].[dbo].[tbl_mailzine_class] "
	sqlStr = sqlStr & " set lastUpdate = getdate(), modiuserid = '' "
	sqlStr = sqlStr & " , itemid1 = " & itemid1 & ", salePer1 = " & salePer1 & ", classDesc1 = '" & classDesc1 & "', classSubDesc1 = '" & classSubDesc1 & "' "
	if (itemid2 <> "") and (itemid3 <> "") then
		sqlStr = sqlStr & " , itemid2 = " & itemid2 & ", salePer2 = " & salePer2 & ", classDesc2 = '" & classDesc2 & "', classSubDesc2 = '" & classSubDesc2 & "' "
		sqlStr = sqlStr & " , itemid3 = " & itemid3 & ", salePer3 = " & salePer3 & ", classDesc3 = '" & classDesc3 & "', classSubDesc3 = '" & classSubDesc3 & "' "
	end if
	sqlStr = sqlStr & " where classDate = '" & classDate & "' "

	''response.write sqlStr&"<br>"
	''response.end
	dbget.execute sqlStr, affectedRows

	response.write "<script type='text/javascript'>"
	if (affectedRows = 1) then
		response.write "	alert('OK');"
	else
		response.write "	alert('----------------------------------------------------------\n\nERROR\n\n----------------------------------------------------------');"
	end if
	response.write "	opener.location.reload();"
	response.write "	opener.focus();"
	response.write "	window.close();"
	response.write "</script>"

elseif (mode = "insmailer") then
	if mailergubun="" or isnull(mailergubun) then
		response.write "메일러 구분이 없습니다."
		dbget.close() : response.end
	end if
	if yyyymmdd="" or isnull(yyyymmdd) then
		response.write "발송일이 없습니다."
		dbget.close() : response.end
	end if

	mailcontents1 	= GetMailzineHtmlMake(idx, "member", mailergubun)
	mailcontents2 	= GetMailzineHtmlMake(idx, "nomember", mailergubun)

	dim filesys, filetxt
	Dim ftp, success, localFilename, remoteFilename
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim objStream, objStreamWithoutBOM

	if (member = "test") then
		if mailergubun="EMS" then
			' POST_ID = ""
			' '// 포스트 아이디 생성
			' sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_GetNewPostIDbyDate] '" & yyyymmdd & "'"
			' rsAppNotiget.Open sqlStr,dbAppNotiget,1
			' if  not rsAppNotiget.EOF  then
			' 	POST_ID = rsAppNotiget("POST_ID")
			' end if
			' rsAppNotiget.Close

			' sqlStr = " select top 1 POST_ID "
			' sqlStr = sqlStr + " from "
			' sqlStr = sqlStr + " [DB_AMailer].[dbo].EMS_MASS_BASE_INFO with (readuncommitted)"
			' sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2018062100001' and JOB_STATUS not in ('99', '40', '41') "
			' rsAppNotiget.Open sqlStr,dbAppNotiget,1
			' if  not rsAppNotiget.EOF  then
			' 	POST_ID = ""
			' end if
			' rsAppNotiget.Close

			' if (POST_ID <> "") then
			' 	sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_AddMassMail] '2018062100001', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "' "
			' 	dbAppNotiget.execute sqlStr, affectedRows

			' 	Set filesys = CreateObject("Scripting.FileSystemObject")
			' 	''Set filetxt = filesys.CreateTextFile(Server.MapPath("html_doc/" + POST_ID + ".html"), True)
			' 	localFilename = Server.MapPath("html_doc/" + POST_ID + ".html")
			' 	localFilename = replace(localFilename,"\mailzine\template\html_doc\","\mailzine\html_doc\")
			' 	remoteFilename = POST_ID + ".html"
			' 	''filetxt.WriteLine(request("mailcontents"))
			' 	''filetxt.Close

				
			' 	Set objStream = Server.CreateObject("ADODB.Stream")

			' 	objStream.Mode = adModeReadWrite
			' 	objStream.Type = 2		' 텍스트 타입 (1: Bin, 2: Text)
			' 	objStream.CharSet = "UTF-8"
			' 	objStream.Open
			' 	objStream.WriteText mailcontents1, 1
			' 	objStream.Position = 3	'Skip BOM bytes

			' 	Set objStreamWithoutBOM = Server.CreateObject("ADODB.Stream")
			' 	objStreamWithoutBOM.Mode = adModeReadWrite
			' 	objStreamWithoutBOM.Type = 1	' 텍스트 타입 (1: Bin, 2: Text)
			' 	objStreamWithoutBOM.Open

			' 	objStream.CopyTo objStreamWithoutBOM

			' 	objStream.Flush
			' 	objStream.Close

			' 	objStreamWithoutBOM.SaveToFile localFilename, 2
			' 	objStreamWithoutBOM.Flush
			' 	objStreamWithoutBOM.Close

			' 	Set objStream = Nothing
			' 	Set objStreamWithoutBOM = Nothing

			' 	set ftp = Server.CreateObject("Chilkat.Ftp2")
			' 	success = ftp.UnlockComponent("10X10CFTP_PzgDkuyF1Yng")

			' 	If (success <> 1) Then
			' 		Response.Write "001. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	ftp.Hostname = "192.168.0.92"
			' 	ftp.Username = "FTP_Amailer"
			' 	ftp.Password = "apdlffj12#"

			' 	'  The default data transfer mode is "Active" as opposed to "Passive".
			' 	''ftp.Passive = 1

			' 	success = ftp.Connect()
			' 	If (success <> 1) Then
			' 		Response.Write "002. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	ftp.PassiveUseHostAddr = 1

			' 	success = ftp.ChangeRemoteDir("/")
			' 	If (success <> 1) Then
			' 		Response.Write "003. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	success = ftp.PutFile(localFilename,remoteFilename)
			' 	If (success <> 1) Then
			' 		Response.Write "004. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If
			' 	ftp.Disconnect

			' 	filesys.DeleteFile(localFilename)

			' end if

		elseif mailergubun="TMS" then
			POST_ID = ""
			'// 포스트 아이디 생성
			sqlStr = " exec tms.[dbo].[usp_TEN_MailerTrans_GetNewPostIDbyDate] '" & yyyymmdd & "'" & vbcrlf

			'response.write sqlStr & "<br>"
			rsTMSget.CursorLocation = adUseClient
			rsTMSget.Open sqlStr, dbTMSget, adOpenForwardOnly, adLockReadOnly
			if  not rsTMSget.EOF  then
				POST_ID = rsTMSget("POST_ID")
			end if
			rsTMSget.Close

			sqlStr = " select top 1 msg_id as POST_ID" & vbcrlf
			sqlStr = sqlStr + " from tms.[dbo].[TMS_CAMP_MSG_INFO] with (readuncommitted)" & vbcrlf
			sqlStr = sqlStr + " where msg_id like '" & Left(POST_ID,8) & "%' and camp_id = '2018062100001' and del_yn='N'" & vbcrlf

			'response.write sqlStr & "<br>"
			rsTMSget.CursorLocation = adUseClient
			rsTMSget.Open sqlStr, dbTMSget, adOpenForwardOnly, adLockReadOnly
			if  not rsTMSget.EOF  then
				POST_ID = ""
			end if
			rsTMSget.Close

			if (POST_ID <> "") then
				sqlStr = " exec [tms].[dbo].[usp_TEN_MailerTrans_AddMassMail] '2018062100001', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "'" & vbcrlf

				'response.write sqlStr & "<br>"
				dbTMSget.execute sqlStr, affectedRows

				Set filesys = CreateObject("Scripting.FileSystemObject")
				''Set filetxt = filesys.CreateTextFile(Server.MapPath("html_doc/" + POST_ID + ".email"), True)
				localFilename = Server.MapPath("html_doc/" + POST_ID + ".email")
				localFilename = replace(localFilename,"\mailzine\template\html_doc\","\mailzine\html_doc\")

				remoteFilename = POST_ID + ".email"
				''filetxt.WriteLine(request("mailcontents"))
				''filetxt.Close

				response.write "localFilename : " & localFilename & "<Br>"
				response.write "remoteFilename : " & remoteFilename & "<Br>"

				Set objStream = Server.CreateObject("ADODB.Stream")

				objStream.Mode = adModeReadWrite
				objStream.Type = 2		' 텍스트 타입 (1: Bin, 2: Text)
				objStream.CharSet = "UTF-8"
				objStream.Open
				objStream.WriteText mailcontents1, 1
				objStream.Position = 3	'Skip BOM bytes

				Set objStreamWithoutBOM = Server.CreateObject("ADODB.Stream")
				objStreamWithoutBOM.Mode = adModeReadWrite
				objStreamWithoutBOM.Type = 1	' 텍스트 타입 (1: Bin, 2: Text)
				objStreamWithoutBOM.Open

				objStream.CopyTo objStreamWithoutBOM

				objStream.Flush
				objStream.Close

				objStreamWithoutBOM.SaveToFile localFilename, 2
				objStreamWithoutBOM.Flush
				objStreamWithoutBOM.Close

				Set objStream = Nothing
				Set objStreamWithoutBOM = Nothing

				' 참고 : https://www.example-code.com/asp/use_explicit_FTP_over_TLS.asp
				set ftp = Server.CreateObject("Chilkat.Ftp2")
				success = ftp.UnlockComponent("10X10CFTP_PzgDkuyF1Yng")
				response.write "UnlockComponent : " & success & "<br>"
				If (success <> 1) Then
					Response.Write "001. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				ftp.Hostname = "192.168.0.110"
				ftp.Username = "ftp_tms"
				ftp.Password = "ftp_tms12#"

				ftp.AuthTls = 1

				'  The default data transfer mode is "Active" as opposed to "Passive".
				''ftp.Passive = 1

				success = ftp.Connect()
				response.write "Connect : " & success & "<br>"
				If (success <> 1) Then
					Response.Write "002. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				ftp.PassiveUseHostAddr = 1

				success = ftp.ChangeRemoteDir("/")
				response.write "ChangeRemoteDir : " & success & "<br>"
				If (success <> 1) Then
					Response.Write "003. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				success = ftp.PutFile(localFilename,remoteFilename)
				response.write "PutFile : " & success & "<br>"
				If (success <> 1) Then
					Response.Write "004. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If
				ftp.Disconnect

				filesys.DeleteFile(localFilename)

			end if
		else
			response.write "정상적인 동작이 아닙니다[0]."
			rsAppNotiget.close() : response.end
		end if
	else
		if mailergubun="EMS" then
			' POST_ID = ""
			' '/////////////// 회원
			' '// 포스트 아이디 생성
			' sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_GetNewPostIDbyDate] '" & yyyymmdd & "'"
			' rsAppNotiget.Open sqlStr,dbAppNotiget,1
			' if  not rsAppNotiget.EOF  then
			' 	POST_ID = rsAppNotiget("POST_ID")
			' end if
			' rsAppNotiget.Close

			' sqlStr = " select top 1 POST_ID "
			' sqlStr = sqlStr + " from "
			' sqlStr = sqlStr + " [DB_AMailer].[dbo].EMS_MASS_BASE_INFO with (readuncommitted)"
			' sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2013032100001' and JOB_STATUS not in ('99', '40', '41') "
			' rsAppNotiget.Open sqlStr,dbAppNotiget,1
			' if  not rsAppNotiget.EOF  then
			' 	POST_ID = ""
			' end if
			' rsAppNotiget.Close

			' if (POST_ID <> "") then
			' 	sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_AddMassMail] '2013032100001', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "' "
			' 	dbAppNotiget.execute sqlStr, affectedRows

			' 	Set filesys = CreateObject("Scripting.FileSystemObject")
			' 	''Set filetxt = filesys.CreateTextFile(Server.MapPath("html_doc/" + POST_ID + ".html"), True)
			' 	localFilename = Server.MapPath("html_doc/" + POST_ID + ".html")
			' 	localFilename = replace(localFilename,"\mailzine\template\html_doc\","\mailzine\html_doc\")
			' 	remoteFilename = POST_ID + ".html"
			' 	''filetxt.WriteLine(request("mailcontents"))
			' 	''filetxt.Close

			' 	Set objStream = Server.CreateObject("ADODB.Stream")

			' 	objStream.Mode = adModeReadWrite
			' 	objStream.Type = 2		' 텍스트 타입 (1: Bin, 2: Text)
			' 	objStream.CharSet = "UTF-8"
			' 	objStream.Open
			' 	objStream.WriteText mailcontents1, 1
			' 	objStream.Position = 3	'Skip BOM bytes

			' 	Set objStreamWithoutBOM = Server.CreateObject("ADODB.Stream")
			' 	objStreamWithoutBOM.Mode = adModeReadWrite
			' 	objStreamWithoutBOM.Type = 1	' 텍스트 타입 (1: Bin, 2: Text)
			' 	objStreamWithoutBOM.Open

			' 	objStream.CopyTo objStreamWithoutBOM

			' 	objStream.Flush
			' 	objStream.Close

			' 	objStreamWithoutBOM.SaveToFile localFilename, 2
			' 	objStreamWithoutBOM.Flush
			' 	objStreamWithoutBOM.Close

			' 	Set objStream = Nothing
			' 	Set objStreamWithoutBOM = Nothing

			' 	set ftp = Server.CreateObject("Chilkat.Ftp2")
			' 	success = ftp.UnlockComponent("10X10CFTP_PzgDkuyF1Yng")

			' 	If (success <> 1) Then
			' 		Response.Write "001. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	ftp.Hostname = "192.168.0.92"
			' 	ftp.Username = "FTP_Amailer"
			' 	ftp.Password = "apdlffj12#"

			' 	'  The default data transfer mode is "Active" as opposed to "Passive".
			' 	''ftp.Passive = 1

			' 	success = ftp.Connect()
			' 	If (success <> 1) Then
			' 		Response.Write "002. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	ftp.PassiveUseHostAddr = 1

			' 	success = ftp.ChangeRemoteDir("/")
			' 	If (success <> 1) Then
			' 		Response.Write "003. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	success = ftp.PutFile(localFilename,remoteFilename)
			' 	If (success <> 1) Then
			' 		Response.Write "004. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If
			' 	ftp.Disconnect

			' 	filesys.DeleteFile(localFilename)

			' end if

			' POST_ID = ""
			' '// 포스트 아이디 생성
			' sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_GetNewPostIDbyDate] '" & yyyymmdd & "'"
			' rsAppNotiget.Open sqlStr,dbAppNotiget,1
			' if  not rsAppNotiget.EOF  then
			' 	POST_ID = rsAppNotiget("POST_ID")
			' end if
			' rsAppNotiget.Close

			' '////////////////// 비회원
			' sqlStr = " select top 1 POST_ID "
			' sqlStr = sqlStr + " from "
			' sqlStr = sqlStr + " [DB_AMailer].[dbo].EMS_MASS_BASE_INFO with (readuncommitted)"
			' ''sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2013032100002' and JOB_STATUS <> '99' "
			' sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2013032100002' and JOB_STATUS not in ('99', '40', '41') "
			' rsAppNotiget.Open sqlStr,dbAppNotiget,1
			' if  not rsAppNotiget.EOF  then
			' 	POST_ID = ""
			' end if
			' rsAppNotiget.Close

			' if (POST_ID <> "") then
			' 	sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_AddMassMail] '2013032100002', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "' "
			' 	dbAppNotiget.execute sqlStr, affectedRows

			' 	Set filesys = CreateObject("Scripting.FileSystemObject")
			' 	''Set filetxt = filesys.CreateTextFile(Server.MapPath("html_doc/" + POST_ID + ".html"), True)
			' 	localFilename = Server.MapPath("html_doc/" + POST_ID + ".html")
			' 	localFilename = replace(localFilename,"\mailzine\template\html_doc\","\mailzine\html_doc\")
			' 	remoteFilename = POST_ID + ".html"
			' 	''filetxt.WriteLine(request("mailcontents"))
			' 	''filetxt.Close

			' 	Set objStream = Server.CreateObject("ADODB.Stream")

			' 	objStream.Mode = adModeReadWrite
			' 	objStream.Type = 2		' 텍스트 타입 (1: Bin, 2: Text)
			' 	objStream.CharSet = "UTF-8"
			' 	objStream.Open
			' 	objStream.WriteText mailcontents2, 1
			' 	objStream.Position = 3	'Skip BOM bytes

			' 	Set objStreamWithoutBOM = Server.CreateObject("ADODB.Stream")
			' 	objStreamWithoutBOM.Mode = adModeReadWrite
			' 	objStreamWithoutBOM.Type = 1	' 텍스트 타입 (1: Bin, 2: Text)
			' 	objStreamWithoutBOM.Open

			' 	objStream.CopyTo objStreamWithoutBOM

			' 	objStream.Flush
			' 	objStream.Close

			' 	objStreamWithoutBOM.SaveToFile localFilename, 2
			' 	objStreamWithoutBOM.Flush
			' 	objStreamWithoutBOM.Close

			' 	Set objStream = Nothing
			' 	Set objStreamWithoutBOM = Nothing

			' 	set ftp = Server.CreateObject("Chilkat.Ftp2")
			' 	success = ftp.UnlockComponent("10X10CFTP_PzgDkuyF1Yng")

			' 	If (success <> 1) Then
			' 		Response.Write "001. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	ftp.Hostname = "192.168.0.92"
			' 	ftp.Username = "FTP_Amailer"
			' 	ftp.Password = "apdlffj12#"

			' 	'  The default data transfer mode is "Active" as opposed to "Passive".
			' 	''ftp.Passive = 1

			' 	success = ftp.Connect()
			' 	If (success <> 1) Then
			' 		Response.Write "002. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	ftp.PassiveUseHostAddr = 1

			' 	success = ftp.ChangeRemoteDir("/")
			' 	If (success <> 1) Then
			' 		Response.Write "003. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If

			' 	success = ftp.PutFile(localFilename,remoteFilename)
			' 	If (success <> 1) Then
			' 		Response.Write "004. <br />"
			' 		Response.Write ftp.LastErrorText & "<br>"
			' 	End If
			' 	ftp.Disconnect

			' 	filesys.DeleteFile(localFilename)

			' end if

		elseif mailergubun="TMS" then
			POST_ID = ""
			'/////////////// 회원
			'// 포스트 아이디 생성
			sqlStr = " exec tms.[dbo].[usp_TEN_MailerTrans_GetNewPostIDbyDate] '" & yyyymmdd & "'" & vbcrlf

			'response.write sqlStr & "<br>"
			rsTMSget.CursorLocation = adUseClient
			rsTMSget.Open sqlStr, dbTMSget, adOpenForwardOnly, adLockReadOnly
			if  not rsTMSget.EOF  then
				POST_ID = rsTMSget("POST_ID")
			end if
			rsTMSget.Close

			sqlStr = " select top 1 msg_id as POST_ID" & vbcrlf
			sqlStr = sqlStr + " from tms.[dbo].[TMS_CAMP_MSG_INFO] with (readuncommitted)" & vbcrlf
			sqlStr = sqlStr + " where msg_id like '" & Left(POST_ID,8) & "%' and camp_id = '2013032100001' and del_yn='N'" & vbcrlf

			'response.write sqlStr & "<br>"
			rsTMSget.CursorLocation = adUseClient
			rsTMSget.Open sqlStr, dbTMSget, adOpenForwardOnly, adLockReadOnly
			if  not rsTMSget.EOF  then
				POST_ID = ""
			end if
			rsTMSget.Close

			if (POST_ID <> "") then
				sqlStr = " exec [tms].[dbo].[usp_TEN_MailerTrans_AddMassMail] '2013032100001', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "'" & vbcrlf

				'response.write sqlStr & "<br>"
				dbTMSget.execute sqlStr, affectedRows

				Set filesys = CreateObject("Scripting.FileSystemObject")
				''Set filetxt = filesys.CreateTextFile(Server.MapPath("html_doc/" + POST_ID + ".email"), True)
				localFilename = Server.MapPath("html_doc/" + POST_ID + ".email")
				localFilename = replace(localFilename,"\mailzine\template\html_doc\","\mailzine\html_doc\")

				remoteFilename = POST_ID + ".email"
				''filetxt.WriteLine(request("mailcontents"))
				''filetxt.Close

				response.write "localFilename : " & localFilename & "<Br>"
				response.write "remoteFilename : " & remoteFilename & "<Br>"

				Set objStream = Server.CreateObject("ADODB.Stream")

				objStream.Mode = adModeReadWrite
				objStream.Type = 2		' 텍스트 타입 (1: Bin, 2: Text)
				objStream.CharSet = "UTF-8"
				objStream.Open
				objStream.WriteText mailcontents1, 1
				objStream.Position = 3	'Skip BOM bytes

				Set objStreamWithoutBOM = Server.CreateObject("ADODB.Stream")
				objStreamWithoutBOM.Mode = adModeReadWrite
				objStreamWithoutBOM.Type = 1	' 텍스트 타입 (1: Bin, 2: Text)
				objStreamWithoutBOM.Open

				objStream.CopyTo objStreamWithoutBOM

				objStream.Flush
				objStream.Close

				objStreamWithoutBOM.SaveToFile localFilename, 2
				objStreamWithoutBOM.Flush
				objStreamWithoutBOM.Close

				Set objStream = Nothing
				Set objStreamWithoutBOM = Nothing

				' 참고 : https://www.example-code.com/asp/use_explicit_FTP_over_TLS.asp
				set ftp = Server.CreateObject("Chilkat.Ftp2")
				success = ftp.UnlockComponent("10X10CFTP_PzgDkuyF1Yng")

				If (success <> 1) Then
					Response.Write "001. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				ftp.Hostname = "192.168.0.110"
				ftp.Username = "ftp_tms"
				ftp.Password = "ftp_tms12#"

				ftp.AuthTls = 1

				'  The default data transfer mode is "Active" as opposed to "Passive".
				''ftp.Passive = 1

				success = ftp.Connect()
				If (success <> 1) Then
					Response.Write "002. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				ftp.PassiveUseHostAddr = 1

				success = ftp.ChangeRemoteDir("/")
				If (success <> 1) Then
					Response.Write "003. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				success = ftp.PutFile(localFilename,remoteFilename)
				If (success <> 1) Then
					Response.Write "004. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If
				ftp.Disconnect

				filesys.DeleteFile(localFilename)

			end if

			POST_ID = ""
			'// 포스트 아이디 생성
			sqlStr = " exec tms.[dbo].[usp_TEN_MailerTrans_GetNewPostIDbyDate] '" & yyyymmdd & "'" & vbcrlf

			'response.write sqlStr & "<br>"
			rsTMSget.CursorLocation = adUseClient
			rsTMSget.Open sqlStr, dbTMSget, adOpenForwardOnly, adLockReadOnly
			if  not rsTMSget.EOF  then
				POST_ID = rsTMSget("POST_ID")
			end if
			rsTMSget.Close

			'////////////////// 비회원
			sqlStr = " select top 1 msg_id as POST_ID" & vbcrlf
			sqlStr = sqlStr + " from tms.[dbo].[TMS_CAMP_MSG_INFO] with (readuncommitted)" & vbcrlf
			sqlStr = sqlStr + " where msg_id like '" & Left(POST_ID,8) & "%' and camp_id = '2013032100002' and del_yn='N'" & vbcrlf

			'response.write sqlStr & "<br>"
			rsTMSget.CursorLocation = adUseClient
			rsTMSget.Open sqlStr, dbTMSget, adOpenForwardOnly, adLockReadOnly
			if  not rsTMSget.EOF  then
				POST_ID = ""
			end if
			rsTMSget.Close

			if (POST_ID <> "") then
				sqlStr = " exec [tms].[dbo].[usp_TEN_MailerTrans_AddMassMail] '2013032100002', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "'" & vbcrlf

				'response.write sqlStr & "<br>"
				dbTMSget.execute sqlStr, affectedRows

				Set filesys = CreateObject("Scripting.FileSystemObject")
				''Set filetxt = filesys.CreateTextFile(Server.MapPath("html_doc/" + POST_ID + ".email"), True)
				localFilename = Server.MapPath("html_doc/" + POST_ID + ".email")
				localFilename = replace(localFilename,"\mailzine\template\html_doc\","\mailzine\html_doc\")

				remoteFilename = POST_ID + ".email"
				''filetxt.WriteLine(request("mailcontents"))
				''filetxt.Close

				response.write "localFilename : " & localFilename & "<Br>"
				response.write "remoteFilename : " & remoteFilename & "<Br>"

				Set objStream = Server.CreateObject("ADODB.Stream")

				objStream.Mode = adModeReadWrite
				objStream.Type = 2		' 텍스트 타입 (1: Bin, 2: Text)
				objStream.CharSet = "UTF-8"
				objStream.Open
				objStream.WriteText mailcontents2, 1
				objStream.Position = 3	'Skip BOM bytes

				Set objStreamWithoutBOM = Server.CreateObject("ADODB.Stream")
				objStreamWithoutBOM.Mode = adModeReadWrite
				objStreamWithoutBOM.Type = 1	' 텍스트 타입 (1: Bin, 2: Text)
				objStreamWithoutBOM.Open

				objStream.CopyTo objStreamWithoutBOM

				objStream.Flush
				objStream.Close

				objStreamWithoutBOM.SaveToFile localFilename, 2
				objStreamWithoutBOM.Flush
				objStreamWithoutBOM.Close

				Set objStream = Nothing
				Set objStreamWithoutBOM = Nothing

				' 참고 : https://www.example-code.com/asp/use_explicit_FTP_over_TLS.asp
				set ftp = Server.CreateObject("Chilkat.Ftp2")
				success = ftp.UnlockComponent("10X10CFTP_PzgDkuyF1Yng")

				If (success <> 1) Then
					Response.Write "001. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				ftp.Hostname = "192.168.0.110"
				ftp.Username = "ftp_tms"
				ftp.Password = "ftp_tms12#"

				ftp.AuthTls = 1

				'  The default data transfer mode is "Active" as opposed to "Passive".
				''ftp.Passive = 1

				success = ftp.Connect()
				If (success <> 1) Then
					Response.Write "002. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				ftp.PassiveUseHostAddr = 1

				success = ftp.ChangeRemoteDir("/")
				If (success <> 1) Then
					Response.Write "003. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If

				success = ftp.PutFile(localFilename,remoteFilename)
				If (success <> 1) Then
					Response.Write "004. <br />"
					Response.Write ftp.LastErrorText & "<br>"
				End If
				ftp.Disconnect

				filesys.DeleteFile(localFilename)

			end if
		else
			response.write "정상적인 동작이 아닙니다[0]."
			rsAppNotiget.close() : response.end
		end if
	end if

	response.write "<script type='text/javascript'>"
	if (POST_ID = "") then
		response.write "	alert('\n\n\n------------------ 등록실패!!! ---------------------\n메일러에 접속해서 등록되어 있는 메일진 삭제 후 등록하세요. \n\n');"
	else
		response.write "	alert('OK');"
	end if

	''response.write "	opener.location.reload();"
	''response.write "	opener.focus();"
	''response.write "	window.close();"
	response.write "</script>"
end if

function GetEventFromMainBanner(yyyymmdd, exc_evt_code, getEa)
	dim sqlStr, result, evtCnt

	evtCnt = 0
	GetEventFromMainBanner = ""

	'// 상단 롤링 배너
	'// 엔조이 기획전
	'// 기획전 모음
	'// 이벤트1~8배너

	'// ========================================================================
	'// 상단 롤링 배너
	'// ========================================================================
	sqlStr = " select top " & getEa & vbcrlf
	sqlStr = sqlStr + "     REPLACE(convert(varchar(200), c.linkurl), '/event/eventmain.asp?eventid=', '') as evt_code " & vbcrlf
	sqlStr = sqlStr + " from " & vbcrlf
	sqlStr = sqlStr + " 	[db_sitemaster].[dbo].tbl_main_contents c with (readuncommitted)" & vbcrlf
	sqlStr = sqlStr + " 	left join [db_sitemaster].[dbo].tbl_main_contents_poscode p with (readuncommitted) on c.poscode=p.poscode " & vbcrlf
	sqlStr = sqlStr + " where " & vbcrlf
	sqlStr = sqlStr + " 	1=1 " & vbcrlf
	sqlStr = sqlStr + " 	and Left(posname,5) <> 'POINT' " & vbcrlf
	sqlStr = sqlStr + " 	and p.gubun = 'index' " & vbcrlf
	sqlStr = sqlStr + " 	and c.enddate>getdate() " & vbcrlf
	sqlStr = sqlStr + " 	and c.isusing='Y'" & vbcrlf
	sqlStr = sqlStr + " 	and c.poscode='710'" & vbcrlf
	sqlStr = sqlStr + " 	and '" & Replace(yyyymmdd, ".", "-") & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) " & vbcrlf
	sqlStr = sqlStr + " 	and c.linkurl like '/event/eventmain.asp?eventid=%' " & vbcrlf
	sqlStr = sqlStr + " 	and REPLACE(convert(varchar(200), c.linkurl), '/event/eventmain.asp?eventid=', '') not in (" & exc_evt_code & ")" & vbcrlf
	sqlStr = sqlStr + " order by c.orderidx asc, c.idx desc"
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		do until rsget.EOF
			if (evtCnt < 12) then
				result = result & "," & rsget("evt_code")
				evtCnt = evtCnt + 1
			end if
			rsget.movenext
		loop
	end if
	rsget.Close

	if (evtCnt >= getEa) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	'// ========================================================================
	'// 엔조이 기획전
	'// ========================================================================
	sqlStr = " select top " & getEa & " c.evt_code "
	sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_enjoy_event] c with (readuncommitted) "
	sqlStr = sqlStr + " where 1=1 and c.StartDate = '" & Replace(yyyymmdd, ".", "-") & "' "
	sqlStr = sqlStr + " order by c.DispOrder asc, c.idx desc "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		do until rsget.EOF
			if (evtCnt < 12) then
				result = result & "," & rsget("evt_code")
				evtCnt = evtCnt + 1
			end if
			rsget.movenext
		loop
	end if
	rsget.Close

	if (evtCnt >= getEa) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	'// ========================================================================
	'// 기획전 모음
	'// ========================================================================
	sqlStr = " select top " & getEa & " convert(varchar, c.Evt_Code1) + ',' + convert(varchar, c.Evt_Code2) + ','"
	sqlStr = sqlStr + " + convert(varchar, c.Evt_Code3) as evt_code "
	sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_gather_event] c with (readuncommitted) "
	sqlStr = sqlStr + " where 1=1 and c.StartDate = '" & Replace(yyyymmdd, ".", "-") & "' "
	sqlStr = sqlStr + " order by c.DispOrder asc, c.idx desc "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		do until rsget.EOF
			if (evtCnt < 12) then
				result = result & "," & rsget("evt_code")
				evtCnt = evtCnt + 3
			end if
			rsget.movenext
		loop
	end if
	rsget.Close

	if (evtCnt >= getEa) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	'// ========================================================================
	'// 이벤트1~8배너
	'// ========================================================================
	sqlStr = " select top " & getEa & " t.eventid as evt_code "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_sitemaster.dbo.tbl_pcmain_enjoyevent t with (readuncommitted) "
	sqlStr = sqlStr + " WHERE 1=1 and isusing='Y' and startdate = '" & Replace(yyyymmdd, ".", "-") & "' and t.eventid not in (" & exc_evt_code & ")"
	sqlStr = sqlStr + " order by t.sortnum asc , t.idx desc "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		do until rsget.EOF
			if (evtCnt < 12) then
				result = result & "," & rsget("evt_code")
				evtCnt = evtCnt + 1
			end if
			rsget.movenext
		loop
	end if
	rsget.Close

	if (evtCnt >= getEa) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	if (result <> "") then
		GetEventFromMainBanner = Mid(result, 2, 1000)
	end if
end function

' MD픽가져오기 		'2019.10.29 정태훈 생성
function GetMDPick(yyyymmdd, excCode, getEa)
	dim sqlStr, result

	GetMDPick = ""

	sqlStr = " select top " & getEa & " f.linkitemid"
	sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_main_mdchoice_flash f with (readuncommitted)"
	sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i with (readuncommitted)"
	sqlStr = sqlStr & " 	on f.linkitemid=i.itemid"
	sqlStr = sqlStr & "		and isnull(i.adultType,0)=0"		' 성인용품제외
	sqlStr = sqlStr & " where f.isusing in ('Y','M')"
	sqlStr = sqlStr & " and convert(varchar(10),f.startdate,120) <= '" & Replace(yyyymmdd, ".", "-") & "'"
	sqlStr = sqlStr & " and convert(varchar(10),f.enddate,120) >= '" & Replace(yyyymmdd, ".", "-") & "'"
	sqlStr = sqlStr & " and f.linkitemid is not NULL"
	sqlStr = sqlStr & " and f.linkitemid not in(" & excCode & ")"
	sqlStr = sqlStr & " order by f.startdate desc,f.disporder ,f.idx desc"

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		do until rsget.EOF
			result = result & "," & rsget("linkitemid")
			rsget.movenext
		loop
	end if
	rsget.Close

	if (result <> "") then
		GetMDPick = Mid(result, 2, 1000)
	end if
end function

' 베스트 가져오기 		'2019.10.29 정태훈 생성
function GetMDPickBest(yyyymmdd, getEa)
	dim sqlStr, resultBest, oaward, ix
	resultBest=""
	set oaward = new SearchItemCls
		oaward.FListDiv 			= "bestlist"
		oaward.FRectSortMethod	    = "be"
		oaward.FPageSize 			= getEa
		oaward.FCurrPage 			= 1
		oaward.FSellScope			= "Y"
		oaward.FScrollCount 		= 1
		oaward.FRectSearchItemDiv   ="D"
		oaward.FminPrice			= 20000
		oaward.FawardType			= "period"
		oaward.fRectadultType=0
		oaward.getSearchList
		If oaward.FResultCount>0 Then
			For ix=0 to oaward.FResultCount-1
				resultBest = resultBest & "," & oaward.FItemList(ix).FItemID
			Next
		end if
	set oaward = Nothing
	if (resultBest <> "") then
		GetMDPickBest = Mid(resultBest, 2, 1000)
	end if
end function

' 신상품 가져오기 		'2019.10.29 정태훈 생성
function GetMDPickNew(yyyymmdd, getEa)
	dim sqlStr, resultNew, oawardNew, ix
	resultNew=""
	set oawardNew = new SearchItemCls
		oawardNew.FListDiv 			= "newlist"
		oawardNew.FRectSortMethod	    = "be"
		oawardNew.FRectSearchFlag 	= "newitem"
		oawardNew.FPageSize 			= getEa
		oawardNew.FCurrPage 			= 1
		oawardNew.FSellScope			= "Y"
		oawardNew.FScrollCount 		= 1
		oawardNew.FRectSearchItemDiv   ="D"
		oawardNew.FminPrice			= 20000
		oawardNew.FSalePercentLow = 0.89
		oawardNew.fRectadultType=0
		oawardNew.getSearchList
		If oawardNew.FResultCount>0 Then
			For ix=0 to oawardNew.FResultCount-1
				resultNew = resultNew & "," & oawardNew.FItemList(ix).FItemID
			Next
		end if
	set oawardNew = Nothing
	if (resultNew <> "") then
		GetMDPickNew = Mid(resultNew, 2, 1000)
	end if
end function

function GetJustOneDay(yyyymmdd)
	dim sqlStr, result

	GetJustOneDay = ""

	sqlStr = " select top 1 j.itemid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_sitemaster].[dbo].tbl_just1day j with (readuncommitted) "
	sqlStr = sqlStr + " where j.JustDate = '" & Replace(yyyymmdd, ".", "-") & "' "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		GetJustOneDay = rsget("itemid")
	end if
	rsget.Close
end function

function GetJustOneDay2018(yyyymmdd, excCode, getEa)
	dim sqlStr, result

	GetJustOneDay2018 = ""

	sqlStr = " select top " & getEa & " idx as itemid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_sitemaster.[dbo].[tbl_just1day2018_list] m with (readuncommitted) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and '" & Replace(yyyymmdd, ".", "-") & "' between m.startdate and m.enddate "
	sqlStr = sqlStr + " 	and m.platform = 'pc' "
	sqlStr = sqlStr + " 	and m.type = 'just1day' "
	sqlStr = sqlStr + " 	and m.isusing = 'Y' "
	sqlStr = sqlStr + " 	and m.idx not in (" & excCode & ")"
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	m.idx desc "
	'response.write sqlStr & "<br>"
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		do until rsget.EOF
			result = result & "," & rsget("itemid")
			rsget.movenext
		loop
	end if
	rsget.Close
	if (result <> "") then
		GetJustOneDay2018 = Mid(result, 2, 1000)
	end if
end function

function GetTenTenClass(yyyymmdd)
	dim sqlStr, result

	GetTenTenClass = ""

	sqlStr = " select top 1 c.itemid1, c.itemid2, c.itemid3 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_sitemaster].[dbo].[tbl_mailzine_class] c with (readuncommitted) "
	sqlStr = sqlStr + " where c.classDate = '" & Replace(yyyymmdd, ".", "-") & "' "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		result = rsget("itemid1")
		if Not IsNull(rsget("itemid2")) then
			result = result & "," & rsget("itemid2")
		end if
		if Not IsNull(rsget("itemid3")) then
			result = result & "," & rsget("itemid3")
		end if
	end if
	rsget.Close
	GetTenTenClass = result
end function
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->
<!-- #include virtual="/lib/db/dbTMSclose.asp" -->