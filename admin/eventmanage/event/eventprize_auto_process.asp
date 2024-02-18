<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_event_winner.asp
' Description :  이벤트 당첨등록
' History : 2007.02.22 정윤정 생성
'           2009.04.14 한용민 수정
'           2009.08.06 허진원 SMS/이메일 발송 추가
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/mailLib.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->

<!-- #include virtual="/lib/util/scm_myalarmlib.asp" -->

<% '쇼핑찬스,한줄낙서,러브하우스,핑거스,위클리코디, 백프로샵,문화이벤트

'--------------------------------------------------------
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim eMode, eCode, egKindCode, ekind ,ename, eKindName, esday, eeday, estate, epday, sType
Dim cEvtCont, strSql, tmpCode, j, iranking, srankname, sgiftname, arrwinner, itemid, stitle, gcd,rg
Dim cvalue, ctype, mprice, csdate, cedate, tlist, cprice, iErrcnt,iSuccnt, iGiftKindCode, dAStartDate , dAEndDate
Dim iEPCode, sGiveWinner, chkSms, smsCont, chkEmail, emailCont, itemuse_sdate, itemuse_edate, usewrite_sdate
Dim usewrite_edate, return_yn, return_date, itemuse_itemid, itemuse, vUploadType, vTempArr, vDelUser, vRemainCount, vSongJangID

'' 배송구분 서동석 추가
Dim isupchebeasong, makerid, reqdeliverdate, PrizeCount
Dim jungsan, jungsanValue, vChangeContents, vSCMChangeSQL

'// MY알림
dim myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL

eMode 		= Request.Form("mode") 	'데이터 처리종류
eCode  		= Request.Form("eC")	'이벤트코드
egKindCode 	= Request.Form("egKC")	'종류별그룹코드(핑거스/문화이벤트 회차)
sType 		= "1" '컬쳐스테이션 고정
vUploadType	= Request.Form("uploadtype")
PrizeCount	= Request.Form("prizecnt")
If vUploadType = "" Then
	vUploadType = "direct"
End If
if egKindCode = "" then egKindCode = 0

IF eCode = 4 THEN
	strSql= " SELECT evt_name "&_
			" FROM [db_culture_station].[dbo].[tbl_culturestation_event] "&_
			" WHERE evt_code ="&egKindCode
	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
		ekind = eCode
		ename = db2html(rsget("evt_name"))
	END IF
	rsget.close
ELSE
set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드

	cEvtCont.fnGetEventCont	 '이벤트 내용 가져오기
	ekind =	cEvtCont.FEKind
	ename =	db2html(cEvtCont.FEName)
	eKindName = fnGetEventCodeDesc("eventkind",ekind)
	esday =	cEvtCont.FESDay
	eeday =	cEvtCont.FEEDay
	epday =	cEvtCont.FEPDay
	estate =	cEvtCont.FEState
set cEvtCont = nothing
END IF


'--------------------------------------------------------
' 데이터 처리  : 이벤트당첨 테이블, 배송, 쿠폰, 각각테이블(한줄,핑거스,러브)
'--------------------------------------------------------

   '기본
	iranking 	= Request.Form("sR")
	srankname 	= html2db(Request.Form("sRN"))

	Dim reqArr, defaltminus, intLoop, ExceptionUser
	If Request.Form("sW") <> "" Then
		reqArr = Request.Form("sW")
		If Right(reqArr,1) = "," Then
			reqArr=left(reqArr,len(reqArr)-1)
		End If
		
		reqArr = Split(reqArr,",")
		PrizeCount = PrizeCount - (ubound(reqArr) + 1)
	Else
		PrizeCount=PrizeCount
	End If

	ExceptionUser = "," & replace(Request.Form("sW"),",","','") & "'"

	'당첨자 자동 등록
	Dim ArrPrizeUser, PrizeUsers
	PrizeUsers=""
	strSQL = "exec [db_culture_station].[dbo].[usp_WWW_CultureStation_AutoPrize_Add] " & PrizeCount & "," & eCode & ",'" & ExceptionUser & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF) then
		ArrPrizeUser = rsget.getRows
	end if
	rsget.Close

	If isArray(ArrPrizeUser) Then
		For intLoop = 0 To UBound(ArrPrizeUser,2)
			PrizeUsers = PrizeUsers + ArrPrizeUser(0,intLoop) + ","
		Next
		PrizeUsers = left(PrizeUsers,len(PrizeUsers)-1)
	Else
		Response.Write "<script>alert('당첨자가 없습니다.');history.back();</script>"
		dbget.close()
		Response.End
	End If

	'2013-01-16 김진영...당첨자ID의 마지막에 ","가 들어가면 이전 페이지로 이동하게 수정
	If Right(PrizeUsers,1) = "," Then
		Call sbAlertMsg ("당첨자의 맨마지막 ,를 빼고 다시입력하세요.\n예를들어 aaa,bbb, -> aaa,bbb", "back", "")
	End If
	'2013-01-16 김진영...당첨자ID의 마지막에 ","가 들어가면 이전 페이지로 이동하게 수정 끝

	arrwinner 	= split(PrizeUsers,",")
	dAStartDate = left(now(),10)
	dAEndDate 	= dateadd("d",14,date())
	If sType = "1" Then
		stitle = html2db(eName & "- " & Trim(Replace(srankname,"당첨","")) & " 당첨") '//OnlyView 일 때 이벤트명+등수별칭
	Else
		stitle 	= html2db(eName&" 당첨") '//이벤트명
	End If
	sGiveWinner =  Request.Form("gUserid")

	'배송
	gcd = "01" '//이벤트:01, 기타:90
	rg = request("rdgubun") '//배송지구분
	iGiftKindCode	= Request.Form("iGK")
	sgiftname		= Request.Form("sGKN")	'//사은품명
	isupchebeasong  =  Request("isupchebeasong")

	jungsan            = request("jungsan")
	jungsanValue       = request("jungsanValue")

	If jungsan = "" Then
		jungsan = "N"
	Else
		jungsan = "Y"
	End If

    makerid         =  Request("makerid")
    reqdeliverdate  =  Request("reqdeliverdate")

	'쿠폰
	cvalue = request("couponvalue")
	ctype = request("coupontype")
	mprice = request("minbuyprice")
	csdate = request("sDate")&" 00:00:00"
	cedate = request("eDate")&" 23:59:59"
	tlist = request("targetitemlist")
	cprice = request("couponmeaipprice")

	'테스터
	itemuse = Replace(request("itemuse"),"'","")
	If sType = "5" Then
		stitle = itemuse
	End If
	itemuse_sdate = request("itemuse_sdate")
	itemuse_edate = request("itemuse_edate")
	usewrite_sdate = request("usewrite_sdate")
	usewrite_edate = request("usewrite_edate")
	return_yn = request("return_yn")
	return_date = request("return_date")
	itemuse_itemid = request("itemuse_itemid")

	'당첨자 메지시
	chkSms = "Y"
	smsCont = "[텐바이텐] 이벤트당첨을 축하합니다. 공지사항 및 마이텐바이텐을 확인해주세요."
	chkEmail = request("chkEmail")
	emailCont = request("emailCont")

	if (Not IsNumeric(cprice)) then cprice=0
	if (cprice="") then cprice=0


	vChangeContents = vChangeContents & "- 이벤트(" & eCode & ") 당첨자 등록." & vbCrLf
	vChangeContents = vChangeContents & "- 구분 = " & sType & ", 등수 = " & iranking & ", 등수별칭 = " & srankname & vbCrLf
	vChangeContents = vChangeContents & "- 당첨확인기간 = " & dAStartDate & " ~ " & dAEndDate & ", 당첨자등록방법 = " & vUploadType & vbCrLf
	vChangeContents = vChangeContents & "- 배송지등록구분 = " & rg & ", 사은품명 = " & sgiftname & ", 출고요청일 = " & reqdeliverdate & vbCrLf
	vChangeContents = vChangeContents & "- 배송구분 = " & isupchebeasong & ", 정산여부 = " & jungsan & ", 정산액 = " & jungsanValue & ", 업체ID = " & makerid & vbCrLf
	vChangeContents = vChangeContents & "- 쿠폰타입 = " & cvalue & "(" & ctype & "), 최소구매금액 = " & mprice & ", 유효기간 = " & csdate & " ~ " & cedate & vbCrLf
	vChangeContents = vChangeContents & "- 테스터상품(" & itemuse_itemid & ") = " & itemuse & ", 테스터상품사용기간 = " & itemuse_sdate & " ~ " & itemuse_edate & vbCrLf
	vChangeContents = vChangeContents & "- 당첨자SMS(" & chkSms & ") = " & smsCont & ", 당첨자이메일(" & chkEmail & ") = " & emailCont & vbCrLf
	vChangeContents = vChangeContents & "- 당첨자 = " & PrizeUsers & vbCrLf

	'####### 당첨자를 엑셀로 업로드.
	If vUploadType = "excel" Then
		strSql = "SELECT Top 100 userid FROM [db_temp].[dbo].[tbl_event_winner_excel] WHERE evt_code = '" & eCode & "'"
		rsget.Open strSql,dbget
		If Not rsget.eof Then
			j = 0
			arrwinner = ""
			Do Until rsget.Eof
				vTempArr = vTempArr & rsget("userid")
				vDelUser = vDelUser & "'" & rsget("userid") & "'"
				
				j = j + 1				
				If rsget.RecordCount <> j Then
					vTempArr = vTempArr & ","
					vDelUser = vDelUser & ","
				End If
				rsget.MoveNext
			Loop
			arrwinner = split(vTempArr,",")
		Else
			Response.Write "<script>alert('엑셀 업로드된 당첨자가 없습니다.');parent.location.reload();</script>"
			dbget.close()
			Response.End
		End If
		rsget.close
	End If
	
	'트랜잭션
	dbget.beginTrans
	IF eMode = "C" THEN
		iEPCode		= request("epC")
		sGiveWinner	= request("gUserid")
	
		'이전 당첨자 상태 변경, 새로운 당첨자 insert
		strSql = "UPDATE  [db_event].[dbo].[tbl_event_prize] SET evtprize_status =6 "&_
				"WHERE evtprize_code= "&iEPCode&" AND evt_winner='"&sGiveWinner&"'"
		dbget.execute strSql
	
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		END IF
	
	END IF
	
	'2013-01-09 김진영 생성	(실제 당첨자ID가 우리 테이블에 있는지 검사// 있으면 진행, 없으면 튕김
	Dim oo
	For oo = 0 to UBound(arrwinner)
		strSql = ""
		strSql = strSql & " SELECT count(*) as cnt FROM db_user.dbo.tbl_user_n where userid='"&html2db(Trim(arrwinner(oo)))&"' " & VBCRLF
		rsget.Open strSql,dbget
			If rsget("cnt") = 0 Then
				rsget.close
				Call sbAlertMsg ("ID : "&html2db(Trim(arrwinner(oo)))&"가 없습니다. 다시 입력하세요", "back", "")
				dbget.close()	:	response.End
			End If
		rsget.Close
	Next
	'2013-01-09 김진영 생성 끝
	
		iErrcnt = 0
	For j = 0 To UBound(arrwinner)
		SELECT CASE eKind
		Case "2" '한줄낙서(이번주 선정자 확인, 5주안에 선정된 사람인지 유무확인) ### 안씀. 사용하려면 맨 아래 주석부분 처리구문 붙여넣기.
		Case "3" '100% shop
		'Case "5" '러브하우스
		Case "8" '디자인핑거스
		Case Else
	
			tmpCode = ""
			vSongJangID = ""
			'1. 이벤트관리 등록
			fnSetEventPrize sType,eCode,egKindCode,iranking,srankname,iGiftKindCode,html2db(Trim(arrwinner(j))),dAStartDate,dAEndDate,session("ssBctId"),iEPCode,stitle
	
			'// MY알림
			myalarmtitle = "이벤트 당첨을 축하드립니다!"
			myalarmsubtitle = ename
			if (Len(myalarmsubtitle) > 20) then
				myalarmsubtitle = Left(myalarmsubtitle, 20) & " ..."
			end if
	
			myalarmcontents = "이벤트 당첨소식을 알려드립니다."
			myalarmwwwTargetURL = "/my10x10/myeventmaster.asp"
	
			Call MyAlarm_InsertMyAlarm_SCM(html2db(Trim(arrwinner(j))), "007", myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL)

			'2. 송장또는 쿠폰 등록
			IF  not( tmpCode = ""  or isNull(tmpCode)) tHEN
				IF CStr(sType) = "3"	THEN '사은품배송
				 	fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname ,iGiftKindCode, isupchebeasong, makerid, reqdeliverdate,jungsanValue, jungsan
				ELSEIF  CStr(sType) ="2" THEN '쿠폰등록
					fnSetUserCoupon html2db(Trim(arrwinner(j))),ctype,cvalue,stitle,mprice,csdate,cedate,tlist,cprice, session("ssBctId"),tmpCode
				ELSEIF  CStr(sType) ="4" THEN '티켓등록
					fnSetTicket tmpCode, egKindCode, html2db(Trim(arrwinner(j)))
				ELSEIF  CStr(sType) ="5" THEN '테스터이벤트
	
					fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname ,0, isupchebeasong, makerid, reqdeliverdate,jungsanValue, jungsan
	
					fnSetTester tmpCode, eCode, html2db(Trim(arrwinner(j))), itemuse_itemid, itemuse, itemuse_sdate, itemuse_edate, usewrite_sdate, usewrite_edate, return_yn, return_date
				END IF
			ELSE
				iErrcnt = iErrcnt + 1
			END IF
	
		END Select
	
		'//당첨자 메시지 발송	### DB 트랜잭션때문에 지체가 되어 발송이 안나가기도 함. 그래서 아래로 옮김. 20151016 강준구.
		'if chkSms="Y" or chkEmail="Y" then
		'	Call fnSendUerMessege(html2db(Trim(arrwinner(j))), chkSms, smsCont, chkEmail, emailCont)
		'end if
		
		
		'### 주소가 없는 사람들은 inputdate 를 null, evtprize_status 를 0으로 바꿔줌.
		If sType = "3" AND rg = "F" Then
			strSql = "IF EXISTS(select id from [db_sitemaster].[dbo].[tbl_etc_songjang] WHERE id = '" & vSongJangID & "' and (reqzipcode = '' or replace(reqzipcode,' ','') = '-')) " & vbCrLf & _
					 "BEGIN " & vbCrLf & _
					 "	UPDATE [db_sitemaster].[dbo].[tbl_etc_songjang] SET inputdate = Null WHERE id = '" & vSongJangID & "' " & vbCrLf & _
					 "	UPDATE [db_event].[dbo].[tbl_event_prize] SET evtprize_status = 0 WHERE evtprize_code = '" & tmpCode & "' " & vbCrLf & _
					 "END "
			dbget.execute strSql
		End If
	Next
	
	If vUploadType = "excel" Then
		strSql = "DELETE FROM [db_temp].[dbo].[tbl_event_winner_excel] WHERE userid IN(" & vDelUser & ") AND evt_code = '" & eCode & "'"
		dbget.execute strSql
		
		strSql = "SELECT count(userid) FROM [db_temp].[dbo].[tbl_event_winner_excel] WHERE evt_code = '" & eCode & "'"
		rsget.Open strSql,dbget
		vRemainCount = rsget(0)
		rsget.close
	End If
	

	IF Err.Number <> 0 THEN
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[10]", "back", "")
	Else
		dbget.CommitTrans
		
		For j = 0 To UBound(arrwinner)
			'//당첨자 메시지 발송
			if chkSms="Y" or chkEmail="Y" then
				Call fnSendUerMessege(html2db(Trim(arrwinner(j))), chkSms, smsCont, chkEmail, emailCont)
			end if
		Next

		'### 수정 로그 저장(event)
		vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & Request("menupos") & "', "
		vSCMChangeSQL = vSCMChangeSQL & "'" & html2db(vChangeContents) & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		dbget.execute(vSCMChangeSQL)
	END IF
	
	If vUploadType = "excel" Then
		Response.Write "<script type=""text/javascript"">" & vbCrLf
		If vRemainCount > 0 Then
			Response.Write "parent.$('#excelprocing').hide();" & vbCrLf
			Response.Write "parent.$('#excelSubmit').show();" & vbCrLf
			Response.Write "parent.$('#excelprocdetail').html('&nbsp;&nbsp;<font color=red size=3>* 남은 처리 갯수 : <strong>"&vRemainCount&"</strong></font>&nbsp;&nbsp;');" & vbCrLf
			Response.Write "parent.jsPageReload();" & vbCrLf
		Else
			Response.Write "alert('등록되었습니다.');" & vbCrLf
			Response.Write "parent.jsPageReload();" & vbCrLf
			Response.Write "parent.window.close();" & vbCrLf
		End If
		Response.Write "</script>" & vbCrLf
	Else
%>
	<script language="javascript">
	<!--
	
		alert("등록되었습니다.");
		opener.location.reload();
		window.close();
	//-->
	</script>
<%
	End If
	
	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 함수생성
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'#### 이벤트 배송등록 #################
	Function fnSetSongjang(ByVal rdgubun, ByVal gubuncd, ByVal gubunname, ByVal evtprize_code, ByVal userid,ByVal prizetitle , ByVal giftkindcode, ByVal isupchebeasong, ByVal makerid, ByVal reqdeliverdate, ByVal jungsanValue, ByVal jungsan)
		if rdgubun="U" then
			strSql = "insert into [db_sitemaster].[dbo].tbl_etc_songjang (gubuncd,gubunname,evtprize_code,userid,prizetitle,evtprize_giftkindcode, isupchebeasong, delivermakerid, reqdeliverdate, jungsan, jungsanYN) "&_
					" values "&_
		 			"  ('" & gubuncd &"','" & gubunname& "',"&evtprize_code &",'"&userid&"','" &prizetitle&"',"&giftkindcode&",'" & isupchebeasong & "','" & makerid & "','" & reqdeliverdate & "','" & jungsanValue & "','" & jungsan & "')"
		elseif rdgubun="F" then
			strSql = "insert into [db_sitemaster].[dbo].tbl_etc_songjang (" & vbcrlf
			strSql = strSql & " gubuncd,gubunname,evtprize_code,userid,username,reqname,reqphone,reqhp,reqzipcode" & vbcrlf
			strSql = strSql & " ,reqaddress1,reqaddress2, inputdate, prizetitle,evtprize_giftkindcode, isupchebeasong" & vbcrlf
			strSql = strSql & " , delivermakerid, reqdeliverdate, jungsan, jungsanYN"
			strSql = strSql & " )" & vbcrlf
			strSql = strSql & " 	select distinct" & vbcrlf
			strSql = strSql & " 	'" & gubuncd & "','" & gubunname & "',"&evtprize_code&", u.userid, u.username, u.username" & vbcrlf
			strSql = strSql & " 	, u.userphone, u.usercell, u.zipcode, u.zipaddr + ' ' + u.useraddr ,u.useraddr, getdate()," & vbcrlf
			strSql = strSql & " 	'" &prizetitle& "',"&giftkindcode&",'" & isupchebeasong & "','" & makerid & "','" & reqdeliverdate & "'" & vbcrlf
			strSql = strSql & " 	,'" & jungsanValue & "','" & jungsan & "'" & vbcrlf
			strSql = strSql & " 	from [db_user].[dbo].tbl_user_n u" & vbcrlf
			strSql = strSql & " 	where u.userid  = '"&userid&"'" & vbcrlf
		end if
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		END IF
		
		strSql = "select @@IDENTITY " '': 작동OK
		rsget.Open strSql, dbget
		vSongJangID = rsget(0)
		rsget.Close
	End Function

	'###쿠폰정보 등록 ###################
	Function fnSetUserCoupon(ByVal userid,ByVal coupontype,ByVal couponvalue,ByVal couponname,ByVal minbuyprice,ByVal startdate,ByVal expiredate,ByVal targetitemlist,ByVal couponmeaipprice,ByVal reguserid, ByVal evtprize_code)
		strSql = "insert into [db_user].[dbo].tbl_user_coupon(masteridx,userid,coupontype,couponvalue,couponname "&_
				 " ,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,reguserid,evtprize_code)"&_
				 " values "&_
				 " (0,'"&userid&"','"&coupontype&"','"&couponvalue&"','"&couponname&"','"&minbuyprice&"',"&_
				 "'"&startdate&"','"&expiredate&"','"&targetitemlist&"',"&couponmeaipprice&",'"&reguserid&"',"&evtprize_code&")"
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		END IF
	END Function

	'###이벤트관리 등록 ###################
	Function fnSetEventPrize(ByVal sType, ByVal eCode,ByVal egKindCode, ByVal evt_ranking,ByVal evt_rankname,ByVal iGiftKindCode,ByVal evt_winner,ByVal dAStartDate,ByVal dAEndDate, ByVal AdminID, ByVal iGiveEPCode,ByVal stitle)
		Dim iprizestatus : iprizestatus = 0
		IF 	(dAEndDate = "" OR (sType="3" and rg="F") )THEN iprizestatus = 3 '당첨확인기간 미 입력시 확인상태로
		IF iGiveEPCode = "" THEN iGiveEPCode = "NULL"
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_prize] (evtprize_type, [evt_code],evtgroup_code, [evt_ranking], [evt_rankname], giftkind_code, [evt_winner],  [evtprize_startdate], [evtprize_enddate], [evtprize_status],[AdminID],[give_evtprizecode],evtprize_name) "&_
				"	 SELECT "&sType&","&eCode&","&egKindCode&","&evt_ranking&",'"&evt_rankname&"','"&iGiftKindCode&"', userid, '"&dAStartDate &"','"&dAEndDate&"', "&iprizestatus&", '"& AdminID&"',"&iGiveEPCode&",'"&stitle&"'"&_
				"		FROM [db_user].[dbo].[tbl_user_n] WHERE userid = '"&evt_winner&"'"


		dbget.execute strSql

		strSql = ""
		'//컬쳐이벤트 당첨자발표 완료 처리 '// 2009-04-14 한용민 당첨자처리
		if eCode = 4 then

			strSql = "update db_culture_station.dbo.tbl_culturestation_event set"+vbcrlf
			strSql = strSql & " prizeyn = 'Y'"+vbcrlf
			strSql = strSql & " where evt_code = "&egKindCode&""+vbcrlf

			'response.write strSql&"<br>"
			dbget.execute strSql

		'//일반이벤트 당첨자발표 완료 처리 '// 2009-04-14 한용민 당첨자처리
		else
			strSql = "update db_event.dbo.tbl_event set"+vbcrlf
			strSql = strSql & " prizeyn = 'Y'"+vbcrlf
			strSql = strSql & " where evt_code = "&eCode&""+vbcrlf

			'response.write strSql&"<br>"
			dbget.execute strSql
		end if

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		END IF

		'' SQL 2005에서는 작동안함..?
		''strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_event_prize] "  '': 작동안함		'/사용금지.전체 라인 몽땅 뿌려짐. '/2016.06.02 한용민
		'strSql = "select SCOPE_IDENTITY()"		'/현재 sql 2005 쓰는곳 없음. 사용할꺼면 이런 형태로 쓸것.	'/2016.06.02 한용민
		''strSql = "select IDENT_CURRENT('[db_event].[dbo].[tbl_event_prize]') " '': 작동OK
		strSql = "select @@IDENTITY " '': 작동OK

		rsget.Open strSql, dbget
		tmpCode = rsget(0)
		rsget.Close
	End Function

	'###티켓 등록 ###################
	Function fnSetTicket(ByVal evtprize_code, ByVal egKindCode, ByVal evt_winner)
		strSql = "INSERT INTO [db_culture_station].[dbo].[tbl_ticket_prize] ( [evtprize_code], [cul_evt_code], [evt_winner])"&_
				" VALUES ("&evtprize_code&","&egKindCode&",'"&evt_winner&"') "
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		END IF
	End Function

	'###테스터 등록 ###################
	Function fnSetTester(ByVal evtprize_code, ByVal evt_Code, ByVal evt_winner, ByVal itemuse_itemid, ByVal itemuse, ByVal itemuse_sdate, ByVal itemuse_edate, ByVal usewrite_sdate, ByVal usewrite_edate, ByVal return_yn, ByVal return_date)
		strSql = "INSERT INTO [db_event].[dbo].[tbl_tester_event_winner] ( [evtprize_code], [evt_code], [evt_winner], [itemid], [itemname], [itemuse_sdate], [itemuse_edate], [usewrite_sdate], [usewrite_edate], [return_yn], [return_date])"&_
				" VALUES ('"&evtprize_code&"','"&evt_Code&"','"&evt_winner&"','"&itemuse_itemid&"','"&itemuse&"','"&itemuse_sdate&"','"&itemuse_edate&"','"&usewrite_sdate&"','"&usewrite_edate&"','"&return_yn&"','"&return_date&"') "
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		END IF
	End Function

	'### SMS/이메일 발송 ###################
	Sub fnSendUerMessege(userid, chks, scont, chke, econt)
		dim uHp, uMail
		strSql = "Select top 1 usercell, usermail " &_
				" From db_user.dbo.tbl_user_n " &_
				" Where userid='" & userid & "'"
		rsget.Open strSql, dbget
		IF Not(rsget.EOF or rsget.BOF) THEN
			uHp = rsget("usercell")
			uMail = rsget("usermail")
		END IF
		rsget.close

		'SMS 발송
		if chks="Y" then
			if Not(uHP="" or isNull(uHP)) then Call SendNormalSMS_LINK(uHP,"",scont)
		end if

		'eMail 발송
		if chke="Y" then
			if Not(uMail="" or isNull(uMail)) then
				Call sendmailCS(uMail, "이벤트 당첨 안내입니다.", replace(econt,vbCrLf,"<br>"))
			end if
		end if
	End Sub
	
'####### 한줄낙서 처리구문
'		'1.Check : 선정여부확인
'		strSql = " SELECT evt_winner FROM  [db_event].[dbo].[tbl_event_prize] WHERE evt_code ="&eCode
'		rsget.Open strSql, dbget
'		IF not (rsget.EOF or rsget.BOF) THEN
'			rsget.close
'			Call sbAlertMsg ("이번주 당첨자가 이미 선정되었습니다", "back", "")
'			dbget.close()	:	response.End
'		END IF
'		rsget.close
'
'		'2.Check : 5주안에 선정된 사람인지 유무확인
'		strSql = " select evt_winner from  [db_event].[dbo].[tbl_event_prize] where evt_code in ( "&_
'				"	select top 5 evt_code from [db_event].[dbo].[tbl_event] where evt_kind = 2  order by evt_code desc "&_
'				")  and evt_winner = '"&html2db(Trim(arrwinner(j)))&"'"
'		rsget.Open strSql, dbget
'		IF not (rsget.EOF or rsget.BOF) THEN
'			rsget.close
'			Call sbAlertMsg ("5주안에 한번이상 당첨되신 분입니다.", "back", "")
'			dbget.close()	:	response.End
'		END IF
'		rsget.close
'
'		'3.Check : 한줄낙서에 글을 쓴 사람인지 유무확인
'		strSql = " select userid from [db_contents].[dbo].[tbl_one_comment]  where  userid='"&html2db(Trim(arrwinner(j)))&"' and evt_code="&eCode
'		rsget.Open strSql, dbget
'		IF (rsget.EOF OR rsget.BOF) THEN
'			rsget.close
'			Call sbAlertMsg ("이벤트 응모자가 아닙니다.당첨자를 확인해주세요", "back", "")
'			dbget.close()	:	response.End
'		END IF
'		rsget.close
'
'		tmpCode = ""
'		'4. 이벤트관리 등록
'		   Call fnSetEventPrize (sType,eCode,egKindCode,iranking,srankname,iGiftKindCode,html2db(Trim(arrwinner(j))),dAStartDate,dAEndDate, session("ssBctId"),iEPCode,stitle)
'
'		   '// MY알림
'		   myalarmtitle = "이벤트 당첨을 축하드립니다!"
'		   myalarmsubtitle = ename
'		   if (Len(myalarmsubtitle) > 20) then
'			   myalarmsubtitle = Left(myalarmsubtitle, 20) & " ..."
'		   end if
'
'		   myalarmcontents = "이벤트 당첨소식을 알려드립니다."
'		   myalarmwwwTargetURL = "/my10x10/myeventmaster.asp"
'
'		   Call MyAlarm_InsertMyAlarm_SCM(html2db(Trim(arrwinner(j))), "007", myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL)
'
'		'5. 송장등록
'		IF  not( tmpCode = ""  or isNull(tmpCode)) tHEN
'			fnSetSongjang rg, gcd, stitle, tmpCode, html2db(Trim(arrwinner(j))), sgiftname , iGiftKindCode, isupchebeasong, makerid, reqdeliverdate,jungsanValue, jungsan
'		ELSE
'			iErrcnt = iErrcnt + 1
'		END IF
'
'		'6.한줄낙서 등록
'		strSql = "UPDATE [db_contents].[dbo].[tbl_one_comment] SET winYN='Y' WHERE userid='"&html2db(Trim(arrwinner(j)))&"' and evt_code="&eCode
'		dbget.execute strSql
'		IF Err.Number <> 0 THEN
'			dbget.RollBackTrans
'			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
'		END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
