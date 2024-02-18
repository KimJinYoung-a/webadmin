<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim i, vQuery, vTemp, vUserID, vCatdID, vCardOpt, vOrderID, vCardPrice, vUserCell, vMMSTitle, vMMSMessage, vIsReg, vMMSOrderID, vMMSDB, vErrQu
	Dim vGiftProcQuery
	vTemp = Trim(Request("userid"))
	vTemp = Replace(vTemp," ","")
	vCatdID = Request("iid")
	vCardOpt = Request("opt")
	vMMSTitle = Request("mmstitle")
	vMMSMessage = Request("mmsmessage")
	
	IF application("Svr_Info") = "Dev" THEN
    	vMMSDB = "[ACADEMYDB].[db_LgSMS].[dbo].[mms_msg]"
    else
    	vMMSDB = "[LOGISTICSDB].[db_LgSMS].[dbo].[mms_msg]"
    end if
	
	'############################################## [1] 카드정보받기 ##################################################
	vQuery = "Select top 1 cardSellCash From [db_item].[dbo].[tbl_giftcard_option] " &_
			" Where cardItemid = '" & vCatdID & "' and cardOption = '" & vCardOpt & "' and optSellYn = 'Y' and optIsUsing = 'Y' "
	rsget.Open vQuery, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		vCardPrice = rsget(0)
	elseif vCardOpt = "0000" then		'수기입력
		vCardPrice = 0
	end if
	rsget.Close
	'############################################## [1] 카드정보받기 ##################################################
	
	
	For i = LBound(Split(vTemp,",")) To UBound(Split(vTemp,","))
	
		vUserID = Split(vTemp,",")(i)
		
		'############################################## [2] 회원정보받기 폰번호 ##################################################
		vQuery = "Select usercell From [db_user].[dbo].[tbl_user_n] Where userid = '" & vUserID & "' "
		rsget.Open vQuery, dbget, 1
		if Not rsget.EOF then
			vUserCell = rsget(0)
			rsget.Close
		else
			rsget.Close
			Response.Write "<script>alert('"&vUserID&" 없는 ID입니다. 확인해보세요.');</script>"
			dbget.close()
			Exit For
			Response.End
		end if
		'############################################## [2] 회원정보받기 폰번호 ##################################################
		
		
		
		'############################################## [3] 주문번호 받아오기 ##################################################
		vOrderID = fnGiftCardReg(vCatdID, vCardOpt, vCardPrice)
		
		If vOrderID = "x" Then
			Response.Write "<script>alert('입력이 안되었습니다. 확인해보세요.');</script>"
			dbget.close()
			Response.End
		End If
		
		vMMSOrderID = vMMSOrderID & "'" & vOrderID & "',"
		'############################################## [3] 주문번호 받아오기 ##################################################
		
		
		
		'############################################## [4] 수령자 휴대폰 및 제목, LMS 내용 적용(등록완료로 수정) ##################################################
		vQuery = "UPDATE [db_order].[dbo].[tbl_giftcard_order] SET " & "<br>" & _
				 "	jumunDiv = '7', sendhp = '1644-6030', reqhp = '" & vUserCell & "', MMSTitle = '" & vMMSTitle & "', " & "<br>" & _
				 "	MMSContent = '" & vMMSMessage & "' " & "<br>" & _
				 " WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		'############################################## [4] 수령자 휴대폰 및 제목, LMS 내용 적용(등록완료로 수정) ##################################################
		
		
		
		'############################################## [5] 등록처리 ##################################################
		vQuery = "INSERT INTO [db_user].[dbo].[tbl_giftcard_regList](giftOrderSerial, masterCardCode, cardItemid, cardOption, cardPrice, buyDate, cardExpire, userid, cardStatus) " & "<br>" & _
				 "	SELECT giftOrderSerial, masterCardCode, cardItemid, cardOption, totalsum, regdate, dateadd(year,5,regdate), '" & vUserID & "', '1' " & "<br>" & _
				 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		'############################################## [5] 등록처리 ##################################################
		
		
		
		'############################################## [6] 로그 추가 ##################################################
		vQuery = "INSERT INTO [db_user].[dbo].[tbl_giftcard_log](userid, useCash, jukyocd, jukyo, orderserial, reguserid, siteDiv) " & "<br>" & _
				 "	SELECT '" & vUserID & "', totalsum, 100, 'GIFT카드 등록', giftOrderSerial, '" & session("ssBctId") & "', 'T' " & "<br>" & _
				 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		'############################################## [6] 로그 추가 ##################################################



		'############################################## [7] 내현황 확인 후 추가 및 수정 ##################################################
		vQuery = "Select distinct userid From [db_user].[dbo].[tbl_giftcard_current] Where userid = '" & vUserID & "' "
		rsget.Open vQuery, dbget, 1
		if Not rsget.EOF then
			vIsReg = "o"
		else
			vIsReg = "x"
		end if
		rsget.close

		If vIsReg = "o" Then	'### 있으면 UPDATE
			vQuery = "UPDATE [db_user].[dbo].[tbl_giftcard_current] SET " & "<br>" & _
					 "	currentCash = (currentCash + " & vCardPrice & "), gainCash = (gainCash + " & vCardPrice & "), lastUpdate = getdate() " & "<br>" & _
					 " WHERE userid = '" & vUserID & "'"
			'dbget.execute vQuery
			vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		ElseIf vIsReg = "x" Then	'### 없으면 INSERT
			vQuery = "INSERT INTO [db_user].[dbo].[tbl_giftcard_current](userid, currentCash, gainCash, lastupdate) " & "<br>" & _
					 "	SELECT '" & vUserID & "', totalsum, totalsum, getdate() " & "<br>" & _
					 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
			'dbget.execute vQuery
			vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		End If
		vErrQu = vErrQu & vQuery & "<br>"
		'############################################## [7] 내현황 확인 후 추가 및 수정 ##################################################
		
		
		
		'############################################## [8] MMS 보내기 ##################################################
		vQuery = "INSERT INTO " & vMMSDB & "(SUBJECT,PHONE,CALLBACK,STATUS,REQDATE,MSG,FILE_CNT, EXPIRETIME) " & "<br>" & _
				 "	SELECT '" & vMMSTitle & "', reqhp, '1644-6030','0',getdate() , '" & vMMSMessage & "','0','43200' " & "<br>" & _
				 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br><br><br><br>"
		'############################################## [8] MMS 보내기 ##################################################
		
		vUserCell = ""
		vIsReg = ""
	Next
	
	vMMSOrderID = Left(vMMSOrderID,Len(vMMSOrderID)-1)
	
	'response.write vErrQu & "<br>"
	'Response.Write "<strong>발급 완료 되었습니다.</strong>"
	Response.write vGiftProcQuery

'===========================================================================================================================================
'####### 필요 함수들 #######
	'### 카드주문번호받기함수
	Function fnGiftCardReg(giftItemid, giftOption, giftcardPrice)
		Dim strSql, rndjumunno, ordUserid, ordUserNm, tmpOrdSn, tmpMstCd, ordIdx
		'### 주문자
		ordUserid = "system"
		'ordUserid = "10x10phone"
		ordUserNm = "텐바이텐"
		
			tmpOrdSn = "": tmpMstCd = ""
		    '임시주문번호 생성
		    Randomize
			rndjumunno = CLng(Rnd * 100000) + 1
			rndjumunno = CStr(rndjumunno)
	
			'@주문건 저장 (GiftCardGbn:0, 추후 1으로 변경;POS수정후)
			strSql = "Insert Into [db_order].[dbo].tbl_giftcard_order "
			strSql = strSql & " (giftOrderSerial,cardItemid,cardOption,masterCardCode,userid,buyname,totalsum,jumundiv,accountdiv,ipkumdiv,ipkumdate "
			strSql = strSql & " ,discountrate,subtotalprice,miletotalprice,tencardspend,referip,userlevel,sumPaymentEtc,designId,resendCnt,GiftCardGbn,notRegSpendSum) "
			strSql = strSql & " Values "
			strSql = strSql & " ('" & rndjumunno & "'," & giftItemid & ",'" & giftOption & "','','" & ordUserid & "','" & ordUserNm & "'," & giftcardPrice
			strSql = strSql & " ,'5','10','8',getdate(),1," & giftcardPrice & ",0,0,'" & Left(request.ServerVariables("REMOTE_ADDR"),32) & "'"
			strSql = strSql & " ,7,0,'101',0,0,0)"
			dbget.Execute strSql
	
			'@IDX접수
			strSql = "Select IDENT_CURRENT('[db_order].[dbo].tbl_giftcard_order') as maxitemid "
			rsget.Open strSql,dbget,1
				ordIdx = rsget("maxitemid")
			rsget.close
	
			'## 실 주문번호/카드코드 Setting
			if (Not IsNull(ordIdx)) and (ordIdx<>"") then
				dim sh: sh = 0
				tmpOrdSn = "G" & Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),4,256)
				tmpOrdSn = tmpOrdSn & Format00(5,Right(CStr(ordIdx),5))
				tmpMstCd = getMasterCode(ordIdx,16,sh)
	
				strSql = " update [db_order].[dbo].tbl_giftcard_order" + vbCrlf
				strSql = strSql + " set giftOrderSerial = '" + tmpOrdSn + "'" + vbCrlf
				strSql = strSql + " ,masterCardCode = '" + tmpMstCd + "'" + vbCrlf
				strSql = strSql + " where idx = " + CStr(ordIdx) + vbCrlf
	
				dbget.Execute strSql
	
				'# 기프트카드 인증번호 발급 로그 저장
				Call putGiftCardMasterCDLog(tmpOrdSn,tmpMstCd,sh-1)
				
				fnGiftCardReg = tmpOrdSn
			else
				fnGiftCardReg = "x"
		    end if
	End Function


    '// 실코드접수(+중복등록검사)
	function getMasterCode(no,sz,byRef sh)
		dim strSql, blChk, bufCode
		blChk = false
		if sh="" then sh=0
		do Until blChk
			if (sz-sh-1)<=0 then blChk=true
			bufCode = makeMasterCode(no,sz,sh)
			strSql = "Select count(idx) from db_order.dbo.tbl_giftcard_cdLog Where masterCardCode='" & bufCode & "'"
			rsget.Open strSql, dbget, 1
				if rsget(0)<=0 then
					IF Not(Left(bufCode,4)="1010" or Left(bufCode,4)="6979") THEN ''preFix 와 중복안되게. (1010: Point1010회원카드, 6979: 실물카드)
					    blChk=true
					    getMasterCode = bufCode
					END IF
				end if
			rsget.Close
			sh = sh +1
		loop
	end function
	
	
	'// 코드생성(일련번호, 코드길이, 중복시프트 / MD5필요)
	function makeMasterCode(no,sz,sh)
		dim tmpMD, tmpNo, tmpChk, i

		'길이 검사
		if (sz>32) or ((31-sz)<sh) then
			makeMasterCode = string(sz,"0")
			exit Function
		end if

		'숫자코드 생성
		tmpMD = MD5(no)
		for i=1 to Len(tmpMD)
			if mid(tmpMD,i,1)>="0" and mid(tmpMD,i,1)<="9" then
				tmpNo = tmpNo & mid(tmpMD,i,1)
			else
				tmpNo = tmpNo & ASC(mid(tmpMD,i,1)) mod 10
			end if
		next

		tmpNo = left(right(tmpNo,len(tmpNo)-sh),sz-1)
		
		'검증코드 생성
		for i=1 to Len(tmpNo)
			tmpChk = tmpChk + (mid(tmpNo,i,1) * i)
		next
		tmpChk = right(tmpChk\Len(tmpNo),1)
		
		'결과 반환
		makeMasterCode = tmpNo & tmpChk
	end function
	
	
	'// 기프트카드 인증번호 발급 로그 저장
	sub putGiftCardMasterCDLog(osn,mcd,sh)
		dim strSql
		strSql = "Insert into db_order.dbo.tbl_giftcard_cdLog " &_
				"(giftOrderSerial, masterCardCode, shiftNum) values " &_
				"('" & osn & "', '" & mcd & "'," & sh & ")"
		dbget.Execute strSql
	end sub
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->