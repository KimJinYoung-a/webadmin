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
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%
dim idx , imgnumber , mode , sql
dim regtype, regdate, evt_code, evtList, itemList
dim title,img1editname,img2editname,img3editname,img4editname,area,gubun,memgubun,secretgubun, just1day, tentenclass
dim classDate, itemid1, salePer1, classDesc1, classSubDesc1, itemid2, salePer2, classDesc2, classSubDesc2, itemid3, salePer3, classDesc3, classSubDesc3
dim sqlStr, affectedRows, member
dim POST_ID, mailcontents, yyyymmdd, mailergubun
	mailergubun = requestcheckvar(request("mailergubun"),16)
	idx 		= requestCheckVar(request("idx"),10)
	imgnumber 	= requestCheckVar(request("imgnumber"),32)
	mode 		= requestCheckVar(request("mode"), 32)
	member 		= requestCheckVar(request("member"), 64)
	mailcontents 	= requestCheckVar(request("mailcontents"), 64000)
	yyyymmdd 	= requestCheckVar(request("yyyymmdd"), 64)

	regtype = requestCheckVar(request("regtype"), 32)
	regdate = requestCheckVar(request("regdate"), 32)
	evt_code = requestCheckVar(request("evt_code"), 32)

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


if mode = "imgdel" then

	if idx = "" or imgnumber = "" then
		response.write "<script>"
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

	response.write "<script>"
	response.write "	alert('OK');"
	response.write "	location.href='/admin/mailzine/mailzine_detail.asp?idx="&idx&"';"
	response.write "</script>"
	dbget.close()	:	response.end

// 가져오기
elseif (mode = "getlist") then
	''response.write regtype & "<br />"
	''response.write regdate & "<br />"
	if (regtype = "2") then
		evtList = GetEventFromMainBanner(regdate, evt_code)
		itemList = GetMDPick(regdate)
		if Replace(regdate, ".", "-") > "2018-09-18" then
			just1day = GetJustOneDay2018(regdate)
		else
			just1day = GetJustOneDay(regdate)
		end if
		tentenclass = GetTenTenClass(regdate)
		response.write "<script>"
		response.write "	parent.document.frm.img1editname.value = '" & Replace(evtList, ",", "\n") & "';"
		response.write "	parent.document.frm.img2editname.value = '" & Replace(itemList, ",", "\n") & "';"
		response.write "	parent.document.frm.img3editname.value = '" & just1day & "';"
		response.write "	parent.document.frm.img4editname.value = '" & Replace(tentenclass, ",", "\n") & "';"
		response.write "	alert('OK');"
		response.write "</script>"
	elseif (regtype = "3") then
		evtList = GetEventFromMainBanner(regdate, evt_code)
		if Replace(regdate, ".", "-") > "2018-09-18" then
			just1day = GetJustOneDay2018(regdate)
		else
			just1day = GetJustOneDay(regdate)
		end if
		tentenclass = GetTenTenClass(regdate)
		response.write "<script>"
		response.write "	parent.document.frm.img1editname.value = '" & Replace(evtList, ",", "\n") & "';"
		response.write "	parent.document.frm.img2editname.value = '';"
		response.write "	parent.document.frm.img3editname.value = '" & just1day & "';"
		response.write "	parent.document.frm.img4editname.value = '" & Replace(tentenclass, ",", "\n") & "';"
		response.write "	alert('OK');"
		response.write "</script>"
	elseif (regtype = "4") then
		evtList = GetEventFromMainBanner(regdate, evt_code)
		itemList = GetMDPick(regdate)
		if Replace(regdate, ".", "-") > "2018-09-18" then
			just1day = GetJustOneDay2018(regdate)
		else
			just1day = GetJustOneDay(regdate)
		end if
		tentenclass = GetTenTenClass(regdate)
		response.write "<script>"
		response.write "	parent.document.frm.img1editname.value = '" & Replace(evtList, ",", "\n") & "';"
		response.write "	parent.document.frm.img2editname.value = '" & Replace(itemList, ",", "\n") & "';"
		response.write "	parent.document.frm.img3editname.value = '" & just1day & "';"
		response.write "	parent.document.frm.img4editname.value = '" & Replace(tentenclass, ",", "\n") & "';"
		response.write "	alert('OK');"
		response.write "</script>"

	' 다이어리스토리
	elseif (regtype = "5") then
		if Replace(regdate, ".", "-") > "2018-09-18" then
			just1day = GetJustOneDay2018(regdate)
		else
			just1day = GetJustOneDay(regdate)
		end if
		tentenclass = GetTenTenClass(regdate)
		response.write "<script type='text/javascript'>"
		response.write "	parent.document.frm.img3editname.value = '" & just1day & "';"
		response.write "	parent.document.frm.img4editname.value = '" & Replace(tentenclass, ",", "\n") & "';"
		response.write "	alert('OK');"
		response.write "</script>"

	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('해당되는 구분이 없습니다.');"
		response.write "</script>"
		dbget.close() : response.end
	end if

' 신규등록
elseif (mode = "ins") then
	if mailergubun="" or isnull(mailergubun) then
		response.write "메일러 구분이 없습니다."
		dbget.close() : response.end
	end if

	if (evt_code = "") then
		evt_code = "NULL"
	end if

	sqlStr = "insert into [db_sitemaster].[dbo].tbl_mailzine" + vbcrlf
	sqlStr = sqlStr & " (title,regdate,imgmap1,imgmap2,imgmap3,imgmap4,area,gubun,memgubun,secretgubun,insertDate, reguserid, regtype, evt_code, mailergubun)" + vbcrlf
	sqlStr = sqlStr & " values(" + vbcrlf
	sqlStr = sqlStr & "'" & html2db(title) & "'," + vbcrlf
	sqlStr = sqlStr & "'" & regdate & "'," + vbcrlf
	sqlStr = sqlStr & "'" & html2db(img1editname) & "'," + vbcrlf
	sqlStr = sqlStr & "'" & html2db(img2editname) & "'," + vbcrlf
	sqlStr = sqlStr & "'" & html2db(img3editname) & "'," + vbcrlf
	sqlStr = sqlStr & "'" & html2db(img4editname) & "'," + vbcrlf
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

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	response.write "<script>"
	response.write "	alert('OK');"
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
	sqlStr = sqlStr & " , imgmap1 = '" & html2db(img1editname) & "' "
	sqlStr = sqlStr & " , imgmap2 = '" & html2db(img2editname) & "' "
	sqlStr = sqlStr & " , imgmap3 = '" & html2db(img3editname) & "' "
	sqlStr = sqlStr & " , imgmap4 = '" & html2db(img4editname) & "' "
	sqlStr = sqlStr & " , area = '" & area & "' "
	sqlStr = sqlStr & " , gubun = '" & gubun & "' "
	sqlStr = sqlStr & " , memgubun = '" & memgubun & "' "
	sqlStr = sqlStr & " , secretgubun = '" & secretgubun & "' "
	sqlStr = sqlStr & " , regtype = '" & regtype & "' "
	sqlStr = sqlStr & " , evt_code = " & evt_code & " "
	sqlStr = sqlStr & " , mailergubun = '" & mailergubun & "' where" & vbcrlf
	sqlStr = sqlStr & " idx = " & idx
	'sqlStr = sqlStr & " 	and reservationDATE is NULL "

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr, affectedRows

	response.write "<script>"
	if (affectedRows = 1) then
		response.write "	alert('OK');"
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

	response.write "<script>"
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

	response.write "<script>"
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
	'// 포스트 아이디 생성
	sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_GetNewPostIDbyDate] '" & yyyymmdd & "'"
	rsAppNotiget.Open sqlStr,dbAppNotiget,1
	if  not rsAppNotiget.EOF  then
		POST_ID = rsAppNotiget("POST_ID")
	end if
	rsAppNotiget.Close

	if (member = "member") then
		sqlStr = " select top 1 POST_ID "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [DB_AMailer].[dbo].EMS_MASS_BASE_INFO "
		sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2013032100001' and JOB_STATUS not in ('99', '40', '41') "
		''sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2018062100001' and JOB_STATUS not in ('99', '40', '41') "
		rsAppNotiget.Open sqlStr,dbAppNotiget,1
		if  not rsAppNotiget.EOF  then
 			POST_ID = ""
		end if
		rsAppNotiget.Close

		if (POST_ID <> "") then
			sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_AddMassMail] '2013032100001', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "' "
			''sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_AddMassMail] '2018062100001', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "' "
			''response.write sqlStr
			dbAppNotiget.execute sqlStr, affectedRows
		end if
	elseif (member = "notmember") then
		sqlStr = " select top 1 POST_ID "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [DB_AMailer].[dbo].EMS_MASS_BASE_INFO "
		''sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2013032100002' and JOB_STATUS <> '99' "
		sqlStr = sqlStr + " where POST_ID like '" & Left(POST_ID,8) & "%' and CAMPAIGN_ID = '2013032100002' and JOB_STATUS not in ('99', '40', '41') "
		rsAppNotiget.Open sqlStr,dbAppNotiget,1
		if  not rsAppNotiget.EOF  then
			POST_ID = ""
		end if
		rsAppNotiget.Close

		if (POST_ID <> "") then
			sqlStr = " exec [DB_AMailer].[dbo].[usp_Ten_MailerTrans_AddMassMail] '2013032100002', '" & POST_ID & "', '" & title & "', '" & yyyymmdd & "' "
			dbAppNotiget.execute sqlStr, affectedRows
		end if
	end if

	dim filesys, filetxt
	Dim ftp, success, localFilename, remoteFilename
	Const ForReading = 1, ForWriting = 2, ForAppending = 8

	if (POST_ID <> "") then
		Set filesys = CreateObject("Scripting.FileSystemObject")
		''Set filetxt = filesys.CreateTextFile(Server.MapPath("html_doc/" + POST_ID + ".html"), True)
		localFilename = Server.MapPath("html_doc/" + POST_ID + ".html")
		remoteFilename = POST_ID + ".html"
		''filetxt.WriteLine(request("mailcontents"))
		''filetxt.Close

		Dim objStream, objStreamWithoutBOM
		Set objStream = Server.CreateObject("ADODB.Stream")

		objStream.Mode = adModeReadWrite
		objStream.Type = 2		' 텍스트 타입 (1: Bin, 2: Text)
		objStream.CharSet = "UTF-8"
		objStream.Open
		objStream.WriteText request("mailcontents"), 1
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

		set ftp = Server.CreateObject("Chilkat.Ftp2")
		success = ftp.UnlockComponent("10X10CFTP_PzgDkuyF1Yng")

		If (success <> 1) Then
			Response.Write "001. <br />"
			Response.Write ftp.LastErrorText & "<br>"
		End If

		ftp.Hostname = "192.168.0.92"
		ftp.Username = "FTP_Amailer"
		ftp.Password = "apdlffj12#"

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

	response.write "<script>"
	if (POST_ID = "") then
		response.write "	alert('\n\n\n------------------ 등록실패!!! ---------------------\n메일러 어드민에 등록되어 있는 메일진 삭제 후 등록하세요. \n\n');"
	else
		response.write "	alert('OK');"
	end if

	''response.write "	opener.location.reload();"
	''response.write "	opener.focus();"
	''response.write "	window.close();"
	response.write "</script>"
end if

function GetEventFromMainBanner(yyyymmdd, exc_evt_code)
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
	sqlStr = " select top 10 "
	sqlStr = sqlStr + "     REPLACE(convert(varchar(200), c.linkurl), '/event/eventmain.asp?eventid=', '') as evt_code "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_sitemaster].[dbo].tbl_main_contents c "
	sqlStr = sqlStr + " 	left join [db_sitemaster].[dbo].tbl_main_contents_poscode p on c.poscode=p.poscode "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1=1 "
	sqlStr = sqlStr + " 	and Left(posname,5) <> 'POINT' "
	sqlStr = sqlStr + " 	and p.gubun = 'index' "
	sqlStr = sqlStr + " 	and c.enddate>getdate() "
	sqlStr = sqlStr + " 	and c.isusing='Y' "
	sqlStr = sqlStr + " 	and c.poscode='710'  "
	sqlStr = sqlStr + " 	and '" & Replace(yyyymmdd, ".", "-") & "' between convert(varchar(10),c.startdate,120) and convert(varchar(10),c.enddate,120) "
	sqlStr = sqlStr + " 	and c.linkurl like '/event/eventmain.asp?eventid=%' "
	sqlStr = sqlStr + " 	and REPLACE(convert(varchar(200), c.linkurl), '/event/eventmain.asp?eventid=', '') <> " & exc_evt_code
	sqlStr = sqlStr + " order by c.orderidx asc, c.idx desc "
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

	if (evtCnt >= 12) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	'// ========================================================================
	'// 엔조이 기획전
	'// ========================================================================
	sqlStr = " select top 10 c.evt_code "
	sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_enjoy_event] c "
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

	if (evtCnt >= 12) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	'// ========================================================================
	'// 기획전 모음
	'// ========================================================================
	sqlStr = " select top 10 convert(varchar, c.Evt_Code1) + ',' + convert(varchar, c.Evt_Code2) + ',' + convert(varchar, c.Evt_Code3) as evt_code "
	sqlStr = sqlStr + " from [db_sitemaster].[dbo].[tbl_main_gather_event] c "
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

	if (evtCnt >= 12) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	'// ========================================================================
	'// 이벤트1~8배너
	'// ========================================================================
	sqlStr = " select top 20 t.eventid as evt_code "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_sitemaster.dbo.tbl_pcmain_enjoyevent t "
	sqlStr = sqlStr + " WHERE 1=1 and isusing='Y' and startdate = '" & Replace(yyyymmdd, ".", "-") & "' and t.eventid <> " & exc_evt_code
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

	if (evtCnt >= 12) then
		if (result <> "") then
			GetEventFromMainBanner = Mid(result, 2, 1000)
		end if
		exit function
	end if

	if (result <> "") then
		GetEventFromMainBanner = Mid(result, 2, 1000)
	end if
end function

function GetMDPick(yyyymmdd)
	dim sqlStr, result

	GetMDPick = ""

	sqlStr = " select top 12 f.linkitemid "
	sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_main_mdchoice_flash f  "
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on f.linkitemid=i.itemid  "
	sqlStr = sqlStr + " where 1=1 and f.isusing in ('Y','M')  "
	sqlStr = sqlStr + " and convert(varchar(10),f.startdate,120) <= '" & Replace(yyyymmdd, ".", "-") & "'  "
	sqlStr = sqlStr + " and convert(varchar(10),f.enddate,120) >= '" & Replace(yyyymmdd, ".", "-") & "'  "
	sqlStr = sqlStr + " and f.linkitemid is not NULL "
	sqlStr = sqlStr + " order by startdate desc,f.disporder ,f.idx desc  "
	rsget.Open sqlStr,dbget,1
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

function GetJustOneDay(yyyymmdd)
	dim sqlStr, result

	GetJustOneDay = ""

	sqlStr = " select top 1 j.itemid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_sitemaster].[dbo].tbl_just1day j "
	sqlStr = sqlStr + " where j.JustDate = '" & Replace(yyyymmdd, ".", "-") & "' "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		GetJustOneDay = rsget("itemid")
	end if
	rsget.Close
end function

function GetJustOneDay2018(yyyymmdd)
	dim sqlStr, result

	GetJustOneDay2018 = ""

	sqlStr = " select top 1 idx as itemid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_sitemaster.[dbo].[tbl_just1day2018_list] m "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and '" & Replace(yyyymmdd, ".", "-") & "' between m.startdate and m.enddate "
	sqlStr = sqlStr + " 	and m.platform = 'pc' "
	sqlStr = sqlStr + " 	and m.type = 'just1day' "
	sqlStr = sqlStr + " 	and m.isusing = 'Y' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	m.idx desc "
	rsget.Open sqlStr,dbget,1
	if  not rsget.EOF  then
		GetJustOneDay2018 = rsget("itemid")
	end if
	rsget.Close
end function

function GetTenTenClass(yyyymmdd)
	dim sqlStr, result

	GetTenTenClass = ""

	sqlStr = " select top 1 c.itemid1, c.itemid2, c.itemid3 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_sitemaster].[dbo].[tbl_mailzine_class] c "
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