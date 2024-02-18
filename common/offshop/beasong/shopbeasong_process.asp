<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->
<%
dim showshopselect, loginidshopormaker

showshopselect = false
loginidshopormaker = ""

if C_ADMIN_USER then
	loginidshopormaker = request("shopid")
elseif (C_IS_SHOP) then
	'직영/가맹점
	loginidshopormaker = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		loginidshopormaker = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'표시안한다. 에러.
		else
			showshopselect = true
			loginidshopormaker = request("shopid")
		end if
	end if
end if

function IsUpcheBeasong(odlvType)
	if (CStr(odlvType) = "2") then
		IsUpcheBeasong = "Y"
	else
		IsUpcheBeasong = "N"
	end if
end function

dim isupchebeasongyn, ExistsBeasongOrderYN, ExistsItemBeasongYN, chkWait, dbCertNo, ordercnt, IpkumDiv
dim i , orderno , itemgubunarr ,itemoptionarr, itemidarr, mode , sql , shopidarr
dim buyname , buyphone1 ,buyphone2 ,buyphone3 ,buyhp1 ,buyhp2 ,buyhp3 ,reqname
dim buyemail1,buyemail2 , reqzipcode,reqzipaddr, reqaddress, comment ,buyphone
dim reqphone1 ,reqphone2 ,reqphone3 ,reqhp1 ,reqhp2 ,reqhp3 , odlvType, tmpcurrstate
dim buyemail ,reqphone ,reqhp ,buyhp ,masteridxtmp , masteridx ,masteridxarr
dim odlvTypearr ,detailidxarr , detailidx, smsyn, KakaoTalkYN, certNo
Dim RectdetailidxArr, RectordernoArr, RectSongjangnoArr, RectSongjangdivArr
dim TotAssignedRow, AssignedRow, FailRow ,ordernoArr, oedit
dim songjangnoArr, songjangdivArr, OrderCount, iMailmasteridxArr, baljunum, baljudate, differencekey
dim mibeasongSoldOutExists, certsendgubun, UserHpAuto, btnJson, minusordernoarr
dim sqlStr, UserHp1, UserHp2, UserHp3, UserHp, smstitlestr, smsmsgstr, kakaomsgstr, RndNo
	UserHp1 = requestcheckvar(request("UserHp1"),4)
	UserHp2 = requestcheckvar(request("UserHp2"),4)
	UserHp3 = requestcheckvar(request("UserHp3"),4)
	UserHp = UserHp1&"-"&UserHp2&"-"&UserHp3
	UserHpAuto = requestcheckvar(request("UserHpAuto"),16)
	odlvType = requestcheckvar(request("odlvType"),1)
	orderno = requestcheckvar(request("orderno"),16)
	itemgubunarr = request("itemgubunarr")
	itemidarr = request("itemidarr")
	itemoptionarr = request("itemoptionarr")
	mode = requestcheckvar(request("mode"),32)
	shopidarr = request("shopidarr")
	buyname = request("buyname")
	reqname = request("reqname")
	buyphone1 = request("buyphone1")
	buyphone2 = request("buyphone2")
	buyphone3 = request("buyphone3")
	buyphone = buyphone1&"-"&buyphone2&"-"&buyphone3
	buyhp1 = request("buyhp1")
	buyhp2 = request("buyhp2")
	buyhp3 = request("buyhp3")
	buyhp = buyhp1&"-"&buyhp2&"-"&buyhp3
	buyemail1 = request("buyemail1")
	buyemail2 = request("buyemail2")
	buyemail = buyemail1&"@"&buyemail2
	reqzipcode = request("reqzipcode")
	reqzipaddr = request("reqzipaddr")
	reqaddress = request("reqaddress")
	comment = request("comment")
	reqphone1 = request("reqphone1")
	reqphone2 = request("reqphone2")
	reqphone3 = request("reqphone3")
	reqphone = reqphone1&"-"&reqphone2&"-"&reqphone3
	reqhp1 = request("reqhp1")
	reqhp2 = request("reqhp2")
	reqhp3 = request("reqhp3")
	reqhp = reqhp1&"-"&reqhp2&"-"&reqhp3
	masteridx =  requestcheckvar(request("masteridx"),10)
	detailidx =  requestcheckvar(request("detailidx"),10)
	masteridxarr = request("masteridxarr")
	odlvTypearr = request("odlvTypearr")
	detailidxarr = request("detailidxarr")
	ordernoArr = request("ordernoArr")
	certsendgubun = requestcheckvar(request("certsendgubun"),32)

ordercnt = 0
ExistsBeasongOrderYN="N"
ExistsItemBeasongYN="N"
smsyn="N"
KakaoTalkYN="N"
chkWait=false

'// 인증번호
Randomize()
RndNo = int(Rnd()*1000000)		'6자리 난수
RndNo = Num2Str(RndNo,6,"0","R")

'// 고객이 직접 주소 입력
if mode = "userjumun" then
	if orderno = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문번호가 없습니다');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	if UserHpAuto = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('휴대폰 번호가 없습니다');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	if certsendgubun = "KAKAOTALK" or certsendgubun = "SMS" then
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문 인증 받으실 구분(카카오톡,SMS)이 없습니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	itemgubunarr = split(itemgubunarr,",")
	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	shopidarr = split(shopidarr,",")
	odlvTypearr = split(odlvTypearr,",")

	if not isarray(shopidarr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('매장 아이디가 없습니다. 관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

'	certNo = md5(trim(orderno) & RndNo & replace(trim(UserHpAuto),"-",""))
'	response.write trim(orderno) & RndNo & replace(trim(UserHpAuto),"-","") & "<Br>"
'	response.write "https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& certNo &""
'	response.end

	sql = "select count(masteridx) as cnt" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master" & vbcrlf
	sql = sql & " where cancelyn='N' and orderno='"& trim(orderno)&"'" & vbcrlf

	'response.write sql &"<br>"
	rsget.open sql ,dbget ,1

	if not(rsget.eof) then
		if rsget("cnt")>0 then
			ExistsBeasongOrderYN = "Y"
		end if
	end if

	rsget.close()

	if ExistsBeasongOrderYN="Y" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('이미 배송이 있는 주문 입니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	dbget.beginTrans

	reqhp = replace(UserHpAuto,"'","")

	if certsendgubun = "KAKAOTALK" then
		KakaoTalkYN="Y"

	elseif certsendgubun = "SMS" then
		smsyn="Y"
	end if

	sql = "update db_shop.dbo.tbl_shopjumun_sms_cert" & vbcrlf
	sql = sql & " set LastUpdate=getdate()" & vbcrlf
	sql = sql & " , isusing='N' where" & vbcrlf
	sql = sql & " isusing='Y' and orderno='"& trim(orderno)&"'" & vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql

	'/ 주문인증정보 등록
	sql = "insert into db_shop.dbo.tbl_shopjumun_sms_cert (shopid, OrderNo, userhp, smsyn, KakaoTalkYN, isusing, Regdate,LastUpdate, CertNo)" & vbcrlf
	sql = sql & " 	select '"& trim(shopidarr(0))&"', '"& trim(orderno)&"', '"&trim(UserHpAuto)&"', '"&smsyn&"', '"&KakaoTalkYN&"', 'Y', getdate(), getdate(), '"& RndNo &"'" & vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql

	'/마스터 테이블 등록
	sql = "insert into" & vbcrlf
	sql = sql & " db_shop.dbo.tbl_shopbeasong_order_master" & vbcrlf
	sql = sql & " (orderno, shopid, ipkumdiv, cancelyn" & vbcrlf
	sql = sql & " ,reqhp,lastupdateadminid) values (" & vbcrlf
	sql = sql & " '"& trim(orderno)&"'" & vbcrlf
	sql = sql & " ,'"& trim(shopidarr(0))&"'" & vbcrlf
	sql = sql & " ,'1'" & vbcrlf
	sql = sql & " ,'N'" & vbcrlf
	sql = sql & " ,'"&trim(reqhp)&"'" & vbcrlf
	sql = sql & " ,'"&session("ssBctId")&"'" & vbcrlf
	sql = sql & " )" & vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql

	masteridxtmp = ""
	sql = ""
	sql = "select max(masteridx) as masteridx" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master" & vbcrlf
	sql = sql & " where cancelyn='N'"

	'response.write sql &"<br>"
	rsget.open sql ,dbget ,1

	if not(rsget.eof) then
		masteridxtmp = rsget("masteridx")
	end if

	rsget.close()

	for i = 0 to ubound(itemgubunarr) - 1

		'//디테일 테이블 등록
		sql = ""
		sql = "insert into" & vbcrlf
		sql = sql & " db_shop.dbo.tbl_shopbeasong_order_detail" & vbcrlf
		sql = sql & " (masteridx, orgdetailidx ,orderno ,itemgubun ,itemid,itemoption" & vbcrlf
		sql = sql & " ,odlvType,isupchebeasong,makerid,itemno,cancelyn,currstate ,lastupdateadminid)" & vbcrlf
		sql = sql & " 	select" & vbcrlf
		sql = sql & " 	'"&masteridxtmp&"', d.idx, m.orderno, d.itemgubun ,d.itemid,d.itemoption" & vbcrlf
		sql = sql & " 	,'"&trim(odlvTypearr(i))&"'" & vbcrlf

		isupchebeasongyn = IsUpcheBeasong(trim(odlvTypearr(i)))
		sql = sql & " ,'" & trim(isupchebeasongyn) & "'" & vbcrlf

		sql = sql & "	,d.makerid ,d.itemno ,'N' ,'0','"&session("ssBctId")&"'" & vbcrlf
		sql = sql & " 	from [db_shop].[dbo].tbl_shopjumun_master m" & vbcrlf
		sql = sql & " 	join [db_shop].[dbo].tbl_shopjumun_detail d" & vbcrlf
		sql = sql & " 	on m.idx = d.masteridx" & vbcrlf
		sql = sql & " 	left join db_shop.dbo.tbl_shopbeasong_order_detail td" & vbcrlf
		sql = sql & " 	on d.idx = td.orgdetailidx and td.cancelyn='N'" & vbcrlf
		sql = sql & " 	where m.cancelyn='N' and d.cancelyn='N'" & vbcrlf
		sql = sql & " 	and m.orderno ='"&trim(orderno)&"'" & vbcrlf
		sql = sql & " 	and td.orderno is null" & vbcrlf	'이미 주문된 내역 제낌
		sql = sql & " 	and d.itemgubun = '"&trim(itemgubunarr(i))&"'" & vbcrlf
		sql = sql & " 	and d.itemid = "&trim(itemidarr(i))&"" & vbcrlf
		sql = sql & " 	and d.itemoption = '"&trim(itemoptionarr(i))&"'" & vbcrlf

		'response.write sql &"<br>"
		dbget.execute sql

	next

	If Err.Number = 0 Then
	    dbget.CommitTrans

		certNo = md5(trim(orderno) & RndNo & replace(trim(UserHpAuto),"-",""))

		smstitlestr = "[텐바이텐] 배송 받으실 주소를 입력해 주세요."
		smsmsgstr = "[텐바이텐] 주문번호: "& trim(orderno) &" 의 주소를 입력해 주세요. " & vbCrLf
		smsmsgstr = smsmsgstr & "https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& certNo &""
		
		btnJson = "{""button"":[{""name"":""주문배송입력/조회"",""type"":""WL"", ""url_mobile"":""https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& certNo &"""}]}"
		kakaomsgstr = "정상적으로 결제 완료되었습니다." & vbCrLf
		kakaomsgstr = kakaomsgstr & "주문해주셔서 감사합니다." & vbCrLf & vbCrLf
		kakaomsgstr = kakaomsgstr & ">주문번호 : " & trim(orderno) & vbCrLf & vbCrLf
		kakaomsgstr = kakaomsgstr & "주문하신 상품에 대한 배송정보 입력은 아래 링크에서 입력해 주시기 바랍니다." & vbCrLf & vbCrLf
		kakaomsgstr = kakaomsgstr & "즐거운 하루 되세요. :D"

		' 카카오톡 발송. 같은 내용을 또 재발송 하면 안됨. IP막힘. 테섭에서도 하지 말것. 실제 발송됨.
		if certsendgubun = "KAKAOTALK" then
			Call SendKakaoMsg_LINK(trim(UserHpAuto),"1644-6030","a-0084",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson)

		' SMS 발송
		elseif certsendgubun = "SMS" then
			sql = "INSERT INTO [SMSDB].db_LgSMS.dbo.MMS_MSG (SUBJECT,PHONE,CALLBACK,STATUS,REQDATE,MSG,FILE_CNT, EXPIRETIME)" & vbcrlf
			sql = sql & " 	select '"& smstitlestr &"', '"& trim(UserHpAuto) &"', '1644-6030','0',getdate(),'"& smsmsgstr &"','0','43200'" & vbcrlf

			'response.write sql &"<br>"
			dbget.execute sql
		end if

		response.write "<script type='text/javascript'>"
		response.write "	alert('주소링크가 고객님께 발송 되었습니다.');"
		response.write "	location.replace('/common/offshop/beasong/shopbeasong_input.asp?orderno="& trim(orderno) &"&menupos="&menupos&"')"
		response.write "</script>"
		dbget.close()	:	response.End

	Else
	    dbget.RollBackTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('값이 일치 하지 않습니다. 관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	End If

'//매장에서 배송요청
elseif mode = "shopjumun" then
	if orderno = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문번호가 없습니다');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	itemgubunarr = split(itemgubunarr,",")
	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	shopidarr = split(shopidarr,",")
	odlvTypearr = split(odlvTypearr,",")

	if not isarray(shopidarr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('매장 아이디가 없습니다. 관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sql = "select count(masteridx) as cnt" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master" & vbcrlf
	sql = sql & " where cancelyn='N' and orderno='"& trim(orderno)&"'" & vbcrlf

	'response.write sql &"<br>"
	rsget.open sql ,dbget ,1

	if not(rsget.eof) then
		if rsget("cnt")>0 then
			ExistsBeasongOrderYN = "Y"
		end if
	end if

	rsget.close()

	if ExistsBeasongOrderYN="Y" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('이미 배송이 있는 주문 입니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	'//코맨트에 뻘짓 해논거 없나 체크
	if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문 유의사항에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "	history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	dbget.beginTrans

	buyemail = replace(buyemail,"'","")
	reqname = replace(reqname,"'","")
	reqzipcode = replace(reqzipcode,"'","")
	reqzipaddr = replace(reqzipaddr,"'","")
	reqaddress = replace(reqaddress,"'","")
	reqphone = replace(reqphone,"'","")
	reqhp = replace(reqhp,"'","")
	comment = replace(comment,"'","""")

	sql = "update db_shop.dbo.tbl_shopjumun_sms_cert" & vbcrlf
	sql = sql & " set LastUpdate=getdate()" & vbcrlf
	sql = sql & " , isusing='N' where" & vbcrlf
	sql = sql & " isusing='Y' and orderno='"& trim(orderno)&"'" & vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql

	'/ 주문인증정보 수기 등록
	sql = "insert into db_shop.dbo.tbl_shopjumun_sms_cert (shopid, OrderNo, userhp, smsyn, KakaoTalkYN, isusing, Regdate,LastUpdate, CertNo)" & vbcrlf
	sql = sql & " 	select '"& trim(shopidarr(0))&"', '"& trim(orderno)&"', '"&trim(reqhp)&"', 'N', 'N', 'Y', getdate(), getdate(), '"& RndNo &"'" & vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql

	'/마스터 테이블 등록
	sql = "insert into" & vbcrlf
	sql = sql & " db_shop.dbo.tbl_shopbeasong_order_master" & vbcrlf
	sql = sql & " (orderno, shopid, ipkumdiv, cancelyn" & vbcrlf		'buyname, buyphone, buyhp
	sql = sql & " , buyemail, reqname, reqzipcode, reqzipaddr, reqaddress, reqphone" & vbcrlf
	sql = sql & " ,reqhp,comment,lastupdateadminid) values (" & vbcrlf
	sql = sql & " '"& trim(orderno) &"'" & vbcrlf
	sql = sql & " ,'"& trim(shopidarr(0))&"'" & vbcrlf
	sql = sql & " ,'2'" & vbcrlf
	sql = sql & " ,'N'" & vbcrlf
	'sql = sql & " ,'"&html2db(trim(buyname))&"'" & vbcrlf
	'sql = sql & " ,'"&trim(buyphone)&"'" & vbcrlf
	'sql = sql & " ,'"&trim(buyhp)&"'" & vbcrlf
	sql = sql & " ,'"&html2db(trim(buyemail))&"'" & vbcrlf
	sql = sql & " ,'"&html2db(trim(reqname))&"'" & vbcrlf
	sql = sql & " ,'"&trim(reqzipcode)&"'" & vbcrlf
	sql = sql & " ,'"&html2db(trim(reqzipaddr))&"'" & vbcrlf
	sql = sql & " ,'"&html2db(trim(reqaddress))&"'" & vbcrlf
	sql = sql & " ,'"&trim(reqphone)&"'" & vbcrlf
	sql = sql & " ,'"&trim(reqhp)&"'" & vbcrlf
	sql = sql & " ,'"&html2db(trim(comment))&"'" & vbcrlf
	sql = sql & " ,'"&session("ssBctId")&"'" & vbcrlf
	sql = sql & " )" & vbcrlf

	'response.write sql &"<br>"
	dbget.execute sql

	masteridxtmp = ""
	sql = ""
	sql = "select max(masteridx) as masteridx" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master" & vbcrlf
	sql = sql & " where cancelyn='N'"

	'response.write sql &"<br>"
	rsget.open sql ,dbget ,1

	if not(rsget.eof) then
		masteridxtmp = rsget("masteridx")
	end if

	rsget.close()

	for i = 0 to ubound(itemgubunarr) - 1

		'//디테일 테이블 등록
		sql = ""
		sql = "insert into" & vbcrlf
		sql = sql & " db_shop.dbo.tbl_shopbeasong_order_detail" & vbcrlf
		sql = sql & " (masteridx, orgdetailidx ,orderno ,itemgubun ,itemid,itemoption" & vbcrlf
		sql = sql & " ,odlvType,isupchebeasong,makerid,itemno,cancelyn,currstate ,lastupdateadminid)" & vbcrlf
		sql = sql & " 	select" & vbcrlf
		sql = sql & " 	'"&masteridxtmp&"', d.idx, m.orderno, d.itemgubun ,d.itemid,d.itemoption" & vbcrlf
		sql = sql & " 	,'"&trim(odlvTypearr(i))&"'" & vbcrlf

		isupchebeasongyn = IsUpcheBeasong(trim(odlvTypearr(i)))
		sql = sql & " ,'" & trim(isupchebeasongyn) & "'" & vbcrlf

		sql = sql & "	,d.makerid ,d.itemno ,'N' ,'0','"&session("ssBctId")&"'" & vbcrlf
		sql = sql & " 	from [db_shop].[dbo].tbl_shopjumun_master m" & vbcrlf
		sql = sql & " 	join [db_shop].[dbo].tbl_shopjumun_detail d" & vbcrlf
		sql = sql & " 	on m.idx = d.masteridx" & vbcrlf
		sql = sql & " 	left join db_shop.dbo.tbl_shopbeasong_order_detail td" & vbcrlf
		sql = sql & " 	on d.idx = td.orgdetailidx and td.cancelyn='N'" & vbcrlf
		sql = sql & " 	where m.cancelyn='N' and d.cancelyn='N'" & vbcrlf
		sql = sql & " 	and m.orderno ='"&trim(orderno)&"'" & vbcrlf
		sql = sql & " 	and td.orderno is null" & vbcrlf	'이미 주문된 내역 제낌
		sql = sql & " 	and d.itemgubun = '"&trim(itemgubunarr(i))&"'" & vbcrlf
		sql = sql & " 	and d.itemid = "&trim(itemidarr(i))&"" & vbcrlf
		sql = sql & " 	and d.itemoption = '"&trim(itemoptionarr(i))&"'" & vbcrlf

		'response.write sql &"<br>"
		dbget.execute sql

	next

	If Err.Number = 0 Then
	    dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('주소 수기 입력이 저장되었습니다.');"
		response.write "	location.replace('/common/offshop/beasong/shopbeasong_input.asp?orderno="& trim(orderno) &"&menupos="&menupos&"')"
		response.write "</script>"
		dbget.close()	:	response.End

	Else
	    dbget.RollBackTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('값이 일치 하지 않습니다. 관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	End If

'//주문 상품 수정
elseif mode = "jumunedit" then
	if orderno = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문 번호가 없습니다');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sql = "select m.orderno" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopjumun_master m" & vbcrlf
	sql = sql & " where m.orderno = '"& Trim(orderno) &"'"
	sql = sql & " and m.cancelyn='N'"

	'response.write sql & "<br>"
	rsget.Open sql,dbget,1
		if not rsget.EOF then
		else
			response.write "<script type='text/javascript'>"
			response.write "	alert('취소되었거나 없는 주문번호 입니다.');"
			response.write "	history.back();"
			response.write "</script>"
			dbget.close()	:	response.End
		end if
	rsget.Close
	
	sql = "select m.masteridx, m.shopid, m.reqhp, m.reqname, m.reqzipcode, m.reqzipaddr, m.reqaddress, m.IpkumDiv" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
	sql = sql & " where m.orderno = '"& Trim(orderno) &"'"
	sql = sql & " and m.cancelyn='N'"

	'response.write sql & "<br>"
	rsget.Open sql,dbget,1
		if not rsget.EOF then
			IpkumDiv = rsget("IpkumDiv")
			masteridx = rsget("masteridx")
		else
			response.write "<script type='text/javascript'>"
			response.write "	alert('배송이 입력이 안된 주문입니다. [OFF]오프_배송관리>>POS_배송입력에서 입력하세요');"
			response.write "	history.back();"
			response.write "</script>"
			dbget.close()	:	response.End
		end if
	rsget.Close

	if IpkumDiv > 5 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('이미 업체에서 확인된 주문입니다.');"
		'response.write "	history.back();"
		response.write "</script>"
	end if
	tmpcurrstate=0
	' 마스터가 배송통보 라면
	if IpkumDiv=5 then
		' 디테일도 통보
		tmpcurrstate = 2
	
	' 아니면 배송대기
	else
		tmpcurrstate = 0
	end if

	detailidxarr = split(detailidxarr,",")
	odlvTypearr = split(odlvTypearr,",")
	itemgubunarr = split(itemgubunarr,",")
	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")

	dbget.beginTrans

	for i = 0 to ubound(itemidarr) - 1
		ExistsItemBeasongYN="N"
	    sql = " select top 1 bd.orderno, bd.currstate"
		sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_detail bd" & vbcrlf
	    sql = sql & " where bd.cancelyn='N'" & vbcrlf
	    sql = sql & " and bd.orderno = '"& trim(orderno) &"'" & vbcrlf
		sql = sql & " and bd.itemgubun = '"&trim(itemgubunarr(i))&"'" & vbcrlf
		sql = sql & " and bd.itemid = "&trim(itemidarr(i))&"" & vbcrlf
		sql = sql & " and bd.itemoption = '"&trim(itemoptionarr(i))&"'" & vbcrlf

		'response.write sql & "<br>"
		'response.end
		rsget.Open sql, dbget, 1

		if Not rsget.Eof then
			ExistsItemBeasongYN="Y"
		end if

		rsget.Close

		' 신규 추가 일경우
		if ExistsItemBeasongYN="N" then
			sql = "insert into db_shop.dbo.tbl_shopbeasong_order_detail (" & vbcrlf
			sql = sql & " masteridx, orgdetailidx ,orderno ,itemgubun ,itemid,itemoption" & vbcrlf
			sql = sql & " ,odlvType,isupchebeasong,makerid,itemno,cancelyn,currstate ,lastupdateadminid)" & vbcrlf
			sql = sql & " 	select" & vbcrlf
			sql = sql & " 	'"& masteridx &"', d.idx, m.orderno, d.itemgubun ,d.itemid,d.itemoption" & vbcrlf
			sql = sql & " 	,'"&trim(odlvTypearr(i))&"'" & vbcrlf

			isupchebeasongyn = IsUpcheBeasong(trim(odlvTypearr(i)))
			sql = sql & " ,'" & trim(isupchebeasongyn) & "'" & vbcrlf

			sql = sql & "	,d.makerid ,d.itemno ,'N' ,'"& tmpcurrstate &"','"&session("ssBctId")&"'" & vbcrlf
			sql = sql & " 	from [db_shop].[dbo].tbl_shopjumun_master m" & vbcrlf
			sql = sql & " 	join [db_shop].[dbo].tbl_shopjumun_detail d" & vbcrlf
			sql = sql & " 		on m.idx = d.masteridx" & vbcrlf
			sql = sql & " 	left join db_shop.dbo.tbl_shopbeasong_order_detail td" & vbcrlf
			sql = sql & " 		on d.idx = td.orgdetailidx and td.cancelyn='N' and td.orderno = '"& trim(orderno) &"'" & vbcrlf
			sql = sql & " 	where m.cancelyn='N' and d.cancelyn='N'" & vbcrlf
			sql = sql & " 	and m.orderno ='"&trim(orderno)&"'" & vbcrlf
			sql = sql & " 	and td.orderno is null" & vbcrlf	'이미 주문된 내역 제낌
			sql = sql & " 	and d.itemgubun = '"&trim(itemgubunarr(i))&"'" & vbcrlf
			sql = sql & " 	and d.itemid = "&trim(itemidarr(i))&"" & vbcrlf
			sql = sql & " 	and d.itemoption = '"&trim(itemoptionarr(i))&"'" & vbcrlf

			'response.write sql &"<br>"
			dbget.execute sql

		else
'			'//디테일 테이블 상태가 주문통보 보다 큰 상품만 빼고 나머지 상품을 수정
'			sql = "update db_shop.dbo.tbl_shopbeasong_order_detail set" & vbcrlf
'			sql = sql & "odlvType = '"& trim(odlvTypearr(i)) &"'" & vbcrlf
'
'			isupchebeasongyn = IsUpcheBeasong(trim(odlvTypearr(i)))
'			sql = sql & " , isupchebeasong = '" & isupchebeasongyn & "'" & vbcrlf
'
'			sql = sql & ",lastupdateadminid = '"&session("ssBctId")&"'" & vbcrlf
'			sql = sql & "from (" & vbcrlf
'			sql = sql & "	select d.detailidx" & vbcrlf
'			sql = sql & "	from db_shop.dbo.tbl_shopbeasong_order_detail d" & vbcrlf
'			sql = sql & "	where d.cancelyn='N'" & vbcrlf
'			sql = sql & "	and d.currstate<=2" & vbcrlf
'			sql = sql & "	and d.orderno = "& trim(orderno) &"" & vbcrlf
'			sql = sql & " 	and d.itemgubun = '"&trim(itemgubunarr(i))&"'" & vbcrlf
'			sql = sql & " 	and d.itemid = "&trim(itemidarr(i))&"" & vbcrlf
'			sql = sql & " 	and d.itemoption = '"&trim(itemoptionarr(i))&"'" & vbcrlf
'			sql = sql & ") as t" & vbcrlf
'			sql = sql & "where db_shop.dbo.tbl_shopbeasong_order_detail.detailidx = t.detailidx" & vbcrlf
'
'			'response.write sql &"<br>"
'			dbget.execute sql
		end if
	next

	' 마스터 일부출고 설정
	sql = "update m set" & vbcrlf
	sql = sql & " m.ipkumdiv = '7'" & vbcrlf
	sql = sql & " , m.beadaldate = getdate()" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
	sql = sql & " join (" & vbcrlf
	sql = sql & " 	select dd.masteridx" & vbcrlf
	sql = sql & " 	from db_shop.dbo.tbl_shopbeasong_order_detail dd" & vbcrlf
	sql = sql & " 	where dd.cancelyn='N'" & vbcrlf
	sql = sql & " 	and dd.currstate < 7" & vbcrlf		' 출고완료 이전인거
	sql = sql & " 	and dd.orderno = '"& trim(orderno) &"'" & vbcrlf
	sql = sql & " 	group by dd.masteridx" & vbcrlf
	sql = sql & " ) as t" & vbcrlf
	sql = sql & " 	on m.masteridx = t.masteridx" & vbcrlf
	sql = sql & " where m.cancelyn='N'" & vbcrlf
	sql = sql & " and m.orderno = '"& trim(orderno) &"'" & vbcrlf

	'response.write sql &"<br>"
    dbget.Execute sql

	' 마스터 출고완료 설정
	sql = "update m set" & vbcrlf
	sql = sql & " m.ipkumdiv = '8'" & vbcrlf
	sql = sql & " , m.beadaldate = getdate()" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
	sql = sql & " left join (" & vbcrlf
	sql = sql & " 	select dd.masteridx" & vbcrlf
	sql = sql & " 	from db_shop.dbo.tbl_shopbeasong_order_detail dd" & vbcrlf
	sql = sql & " 	where dd.cancelyn='N'" & vbcrlf
	sql = sql & " 	and dd.currstate < 7" & vbcrlf		' 출고완료 이전인거
	sql = sql & " 	and dd.orderno = '"& trim(orderno) &"'" & vbcrlf
	sql = sql & " 	group by dd.masteridx" & vbcrlf
	sql = sql & " ) as t" & vbcrlf
	sql = sql & " 	on m.masteridx = t.masteridx" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " where m.cancelyn='N'" & vbcrlf
	sql = sql & " and m.orderno = '"& trim(orderno) &"'" & vbcrlf
	sql = sql & " and t.masteridx is null"

	'response.write sql &"<br>"
    dbget.Execute sql

	If Err.Number = 0 Then
	    dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	location.href='/common/offshop/beasong/shopbeasong_input.asp?orderno="& orderno &"&menupos="& menupos &"';"
		response.write "	alert('처리되었습니다');"
		response.write "</script>"
		dbget.close()	:	response.End

	Else
	    dbget.RollBackTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('값이 일치 하지 않습니다. 관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	End If

' 주문인증정보 수정 , 카카오톡 재발송, sms 재발송
elseif mode="certedit" or mode="ReSendKakaotalk" or mode="ReSendSMS" then
	if orderno = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문번호가 없습니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	set oedit = new cupchebeasong_list
		oedit.frectorderno = orderno
		oedit.fshopjumun_edit()

	if oedit.ftotalcount > 0 then
		dbCertNo = oedit.FOneItem.fCertNo
	end if

	set oedit = nothing

	UserHp = replace(UserHp,"'","")

	certNo = md5(trim(orderno) & dbCertNo & replace(trim(UserHp),"-",""))

'	response.write trim(orderno) & dbCertNo & replace(trim(UserHp),"-","") & "<Br>"
'	response.write "https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& certNo &""
'	response.end

	smstitlestr = "[텐바이텐] 배송 받으실 주소를 입력해 주세요."
	smsmsgstr = "[텐바이텐] 주문번호: "& trim(orderno) &" 의 주소를 입력해 주세요. " & vbCrLf
	smsmsgstr = smsmsgstr & "https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& certNo &""
	
	btnJson = "{""button"":[{""name"":""주문배송입력/조회"",""type"":""WL"", ""url_mobile"":""https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& certNo &"""}]}"
	kakaomsgstr = "정상적으로 결제 완료되었습니다." & vbCrLf
	kakaomsgstr = kakaomsgstr & "주문해주셔서 감사합니다." & vbCrLf & vbCrLf
	kakaomsgstr = kakaomsgstr & ">주문번호 : " & trim(orderno) & vbCrLf & vbCrLf
	kakaomsgstr = kakaomsgstr & "주문하신 상품에 대한 배송정보 입력은 아래 링크에서 입력해 주시기 바랍니다." & vbCrLf & vbCrLf
	kakaomsgstr = kakaomsgstr & "즐거운 하루 되세요. :D"

	' 카카오톡 발송. 같은 내용을 또 재발송 하면 안됨. IP막힘. 테섭에서도 하지 말것. 실제 발송됨.
	if mode="ReSendKakaotalk" then
		sql = "select count(authidx)" & vbcrlf
		sql = sql & " from db_shop.dbo.tbl_shopjumun_sms_cert" & vbcrlf
		sql = sql & " where OrderNo='"& trim(orderno)&"'" & vbcrlf
		sql = sql & " and datediff(ss, isnull(LastUpdate,Regdate) ,getdate()) between 0 and 180" & vbcrlf
		sql = sql & " and isusing='Y'"

		'response.write sql & "<br>"
		rsget.Open sql,dbget,1
			chkWait = rsget(0)>0
		rsget.Close

		if chkWait then
			response.write "<script type='text/javascript'>"
			response.write "	alert('이미 고객님께 주소 입력 링크가 발송 되었습니다. 3분후에 이용 가능 합니다.');"
			response.write "	history.back();"
			response.write "</script>"
			dbget.close()	:	response.End
		end if	

		Call SendKakaoMsg_LINK(trim(UserHp),"1644-6030","a-0084",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson)

		sql = "update db_shop.dbo.tbl_shopjumun_sms_cert" & vbcrlf
		sql = sql & " set KakaoTalkYN='Y'" & vbcrlf
		sql = sql & " , LastUpdate=getdate() where" & vbcrlf
		sql = sql & " isusing='Y' and orderno = '"&trim(orderno)&"'"

		'response.write sql &"<br>"
		dbget.execute sql

	' sms 발송 일경우
	elseif mode="ReSendSMS" then
		' SMS 발송
		sql = "INSERT INTO [SMSDB].db_LgSMS.dbo.MMS_MSG (SUBJECT,PHONE,CALLBACK,STATUS,REQDATE,MSG,FILE_CNT, EXPIRETIME)" & vbcrlf
		sql = sql & " 	select '"& smstitlestr &"', '"& trim(UserHp) &"', '1644-6030','0',getdate(),'"& smsmsgstr &"','0','43200'" & vbcrlf

		'response.write sql &"<br>"
		dbget.execute sql

		sql = "update db_shop.dbo.tbl_shopjumun_sms_cert" & vbcrlf
		sql = sql & " set smsyn='Y'" & vbcrlf
		sql = sql & " , LastUpdate=getdate() where" & vbcrlf
		sql = sql & " isusing='Y' and orderno = '"&trim(orderno)&"'"

		'response.write sql &"<br>"
		dbget.execute sql
	end if

	sql = "update db_shop.dbo.tbl_shopjumun_sms_cert" & vbcrlf
	sql = sql & " set LastUpdate=getdate()" & vbcrlf
	sql = sql & " , UserHp = '"& trim(UserHp) &"' where" & vbcrlf
	sql = sql & " isusing='Y' and orderno = '"&trim(orderno)&"'"

	'response.write sql &"<br>"
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('처리되었습니다');"
	response.write "	location.href='/common/offshop/beasong/shopbeasong_input.asp?orderno="& orderno &"&menupos="& menupos &"';"
	response.write "</script>"
	dbget.close()	:	response.End

'//배송지 정보 수정
elseif mode="addressedit" then
	if masteridx = "" and orderno = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('마스터["& masteridx &"] 값이나 주문번호["& orderno &"] 둘중에 하나의 값은 있어야 합니다.');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	'//코맨트에 뻘짓 해논거 없나 체크
	if checkNotValidHTML(comment) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문 유의사항에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "	history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	buyemail = replace(buyemail,"'","")
	reqname = replace(reqname,"'","")
	reqzipcode = replace(reqzipcode,"'","")
	reqzipaddr = replace(reqzipaddr,"'","")
	reqaddress = replace(reqaddress,"'","")
	reqphone = replace(reqphone,"'","")
	reqhp = replace(reqhp,"'","")
	comment = replace(comment,"'","""")

	sql = "update db_shop.dbo.tbl_shopbeasong_order_master" & vbcrlf
	sql = sql & " set ipkumdiv='2'" & vbcrlf
	sql = sql & " , buyemail = '"&html2db(trim(buyemail))&"'" & vbcrlf
	sql = sql & " ,reqname = '"&html2db(trim(reqname))&"'" & vbcrlf
	sql = sql & " ,reqzipcode = '"&trim(reqzipcode)&"'" & vbcrlf
	sql = sql & " ,reqzipaddr = '"&html2db(trim(reqzipaddr))&"'" & vbcrlf
	sql = sql & " ,reqaddress = '"&html2db(trim(reqaddress))&"'" & vbcrlf
	sql = sql & " ,reqphone = '"&trim(reqphone)&"'" & vbcrlf
	sql = sql & " ,reqhp = '"&trim(reqhp)&"'" & vbcrlf
	sql = sql & " ,comment = '"&html2db(trim(comment))&"'" & vbcrlf
	sql = sql & " ,lastupdateadminid = '"&session("ssBctId")&"' where" & vbcrlf

	if masteridx<>"" then
		sql = sql & " masteridx = "&trim(masteridx)&""
	elseif orderno <> "" then
		sql = sql & " orderno = '"&trim(orderno)&"'"
	end if

	'response.write sql &"<br>"
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('저장 되었습니다.');"
	response.write "	location.href='/common/offshop/beasong/shopbeasong_input.asp?orderno="& orderno &"&menupos="& menupos &"';"
	response.write "</script>"
	dbget.close()	:	response.End

'//배송통보
elseif mode="beasonginput" then
	if ordernoarr = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('주문번호가 없습니다.[0]');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	ordernoarr = split(ordernoarr,",")

	dbget.beginTrans

	for i = 0 to ubound(ordernoarr) - 1

		if trim(ordernoarr(i)) = "" then
			dbget.RollBackTrans
			response.write "<script type='text/javascript'>"
			response.write "	alert('주문번호가 없습니다.[1]');"
			response.write "	history.back();"
			response.write "</script>"
			response.End
		end if

		sql = "select m.shopid, m.reqhp, m.reqname, m.reqzipcode, m.reqzipaddr, m.reqaddress" & vbcrlf
		sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
		sql = sql & " join db_shop.dbo.tbl_shopjumun_master om" & vbcrlf
		sql = sql & " 	on m.orderno = om.orderno" & vbcrlf
		sql = sql & " 	and om.cancelyn='N'" & vbcrlf
		sql = sql & " where m.orderno = '"& trim(ordernoarr(i)) &"'"
		sql = sql & " and m.cancelyn='N'" & vbcrlf

		'response.write sql & "<br>"
		rsget.Open sql,dbget,1
			if not rsget.EOF then
				reqname = rsget("reqname")
				reqzipcode = rsget("reqzipcode")
				reqzipaddr = rsget("reqzipaddr")
				shopidarr = rsget("shopid") & ","
			else
				dbget.RollBackTrans
				response.write "<script type='text/javascript'>"
				response.write "	alert('정상적인 주문이 아닙니다.(주문번호 : "& trim(ordernoarr(i)) &")');"
				response.write "	history.back();"
				response.write "</script>"
				rsget.Close : response.End
			end if
		rsget.Close
		shopidarr = split(shopidarr,",")

	    '마이너스 주문이 있는지 확인
	    sql = " select distinct m.orderno"
		sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
	    sql = sql & " join db_shop.dbo.tbl_shopbeasong_order_detail d" & vbcrlf
	    sql = sql & " 	on m.masteridx = d.masteridx" & vbcrlf
	    sql = sql & " where m.cancelyn='N' and d.cancelyn='N'" & vbcrlf
	    sql = sql & " and m.ipkumdiv<5" & vbcrlf
	    sql = sql & " and d.currstate<2" & vbcrlf
	    sql = sql & " and m.orderno = '"& trim(ordernoarr(i)) &"'" & vbcrlf
	    sql = sql & " and d.itemno < 0 "

		'response.write sql & "<br>"
		'response.end
		rsget.Open sql, dbget, 1

		if Not rsget.Eof then
			do until rsget.eof
				if (minusordernoarr = "") then
					minusordernoarr = rsget("orderno")
				else
					minusordernoarr = minusordernoarr + "," + rsget("orderno")
				end if
				rsget.movenext
			loop
		end if
		rsget.close

		if (minusordernoarr <> "") then
			response.write "<script type='text/javascript'>"
			response.write "	alert('주문중에 마이너스 주문이 있는 주문(" & minusordernoarr & ")이 있습니다.');"
			response.write "	history.back();"
			response.write "</script>"
			response.End
		end if

		if reqname="" then
			dbget.RollBackTrans
			response.write "<script type='text/javascript'>"
			response.write "	alert('배송 받으실분 성함이 없습니다.(주문번호 : "& trim(ordernoarr(i)) &")');"
			response.write "	history.back();"
			response.write "</script>"
			dbget.close()	: response.End
		end if
		if reqzipcode="" or reqzipaddr="" then
			dbget.RollBackTrans
			response.write "<script type='text/javascript'>"
			response.write "	alert('배송 받으실 주소가 없습니다.(주문번호 : "& trim(ordernoarr(i)) &")');"
			response.write "	history.back();"
			response.write "</script>"
			dbget.close()	: response.End
		end if

		'//디테일 테이블 상태가 주문통보 보다 큰 상품만 빼고 나머지 상품을 주문통보 상태로 바꾼다
		sql = "update db_shop.dbo.tbl_shopbeasong_order_detail" & vbcrlf
		sql = sql & " set currstate = (case when odlvType = '0' then '3' else '2' end)" & vbcrlf	'//매장배송일 경우 바로 주문확인 상태로
		sql = sql & " , lastupdateadminid = '"&session("ssBctId")&"' where " & vbcrlf
		sql = sql & " cancelyn='N'" & vbcrlf
		sql = sql & " and currstate<2" & vbcrlf
		sql = sql & " and orderno = '"& trim(ordernoarr(i)) &"'"

		'response.write sql &"<br>"
		'response.end
		dbget.execute sql

		'//디테일 테이블 상태가 배송통보 상태 보다 큰 상품이 존재 하지 않으면 마스터 테이블 상태를 배송통보로 바꾼다
		sql = "update m set" & vbcrlf
		sql = sql & " m.ipkumdiv = '5'" & vbcrlf
		sql = sql & " , m.baljudate = getdate()" & vbcrlf
		sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
		sql = sql & " join (" & vbcrlf
		sql = sql & " 	select dd.masteridx" & vbcrlf
		sql = sql & " 	from db_shop.dbo.tbl_shopbeasong_order_detail dd" & vbcrlf
		sql = sql & " 	where dd.cancelyn='N'" & vbcrlf
		sql = sql & " 	and dd.currstate < 4" & vbcrlf
		sql = sql & " 	and dd.orderno = '"& trim(ordernoarr(i)) &"'" & vbcrlf
		sql = sql & " 	group by dd.masteridx" & vbcrlf
		sql = sql & " ) as t" & vbcrlf
		sql = sql & " 	on m.masteridx = t.masteridx" & vbcrlf
		sql = sql & " where m.cancelyn='N'" & vbcrlf
		sql = sql & " and m.ipkumdiv < 5" & vbcrlf
		sql = sql & " and m.orderno = '"& trim(ordernoarr(i)) &"'"

		'response.write sql &"<br>"
		dbget.execute sql

		' 물류 출고단 기타출고에 꽂는다.
		sql = "insert into [db_sitemaster].[dbo].tbl_etc_songjang (" & vbcrlf
		sql = sql & " gubuncd, gubunname, userid, username, reqname, reqphone, reqhp, reqzipcode, reqaddress1" & vbcrlf
		sql = sql & " , reqaddress2, reqetc, inputdate, isupchebeasong, reqdeliverdate, etckey" & vbcrlf
		sql = sql & " )" & vbcrlf
		sql = sql & " 	select" & vbcrlf
		sql = sql & " 	'70', '매장출고 '+m.orderno, '', m.reqname, m.reqname, m.reqphone, m.reqhp, m.reqzipcode" & vbcrlf
		sql = sql & " 	, m.reqzipaddr, m.reqaddress, m.comment, getdate(), 'N'" & vbcrlf
		sql = sql & " 	, convert(varchar(10),getdate(),21), m.orderno" & vbcrlf
		sql = sql & " 	from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
		sql = sql & " 	join db_shop.dbo.tbl_shopbeasong_order_detail d" & vbcrlf
		sql = sql & " 		on m.masteridx = d.masteridx" & vbcrlf
		sql = sql & " 		and d.odlvType = '1'" & vbcrlf
		sql = sql & " 		and d.isupchebeasong='N'" & vbcrlf
		sql = sql & " 		and d.cancelyn='N'" & vbcrlf
		sql = sql & " 		and d.currstate>=2" & vbcrlf
		sql = sql & " 	left join [db_sitemaster].[dbo].tbl_etc_songjang w" & vbcrlf
		sql = sql & " 		on m.orderno = w.etckey" & vbcrlf
		sql = sql & " 		and w.deleteyn='N'" & vbcrlf
		sql = sql & " 		and w.issended='Y'" & vbcrlf
		sql = sql & " 	where m.cancelyn='N'" & vbcrlf
		sql = sql & " 	and m.cancelyn='N'" & vbcrlf
		sql = sql & " 	and m.ipkumdiv = '5'" & vbcrlf
		sql = sql & " 	and w.etckey is null" & vbcrlf		' 이미 보낸거 제낌
		sql = sql & " 	and m.orderno = '"& trim(ordernoarr(i)) &"'" & vbcrlf
		sql = sql & " 	group by '매장출고 '+m.orderno, m.reqname, m.reqname, m.reqphone, m.reqhp, m.reqzipcode" & vbcrlf
		sql = sql & " 	, m.reqzipaddr, m.reqaddress, m.comment, m.orderno, m.masteridx" & vbcrlf
		sql = sql & " 	order by m.masteridx asc" & vbcrlf

		'response.write sql &"<br>"
		dbget.execute sql
	next

	If Err.Number = 0 Then
	    dbget.CommitTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('처리되었습니다');"
		response.write "	location.replace('/common/offshop/beasong/shopbeasong_list.asp?menupos="& menupos &"');"
		response.write "</script>"
		dbget.close()	:	response.End

	Else
	    dbget.RollBackTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('값이 일치 하지 않습니다. 관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	End If

'//디테일 상품 삭제
elseif mode="detaildel" then

	if detailidx = "" or orderno = "" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('값이 없습니다.관리자 문의하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	dbget.beginTrans

	sql = "update db_shop.dbo.tbl_shopbeasong_order_detail set" & vbcrlf
	sql = sql & " cancelyn='Y'" & vbcrlf
	sql = sql & " ,lastupdateadminid = '"&session("ssBctId")&"'" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_detail" & vbcrlf
	sql = sql & " where cancelyn='N'" & vbcrlf
	sql = sql & " and detailidx = "&detailidx&" and orderno = '"& orderno &"'" & vbcrlf

	'//매장 배송일 경우 출고 완료 상태인것도 삭제하고 다시 배송입력 가능
	if odlvType="0" then

		sql = sql & " and currstate<>7"

	'//물류 배송과 업체배송일 경우 주문 확인 이전 내역만 삭제
	'업체페이지에서 출고완료내역까지 삭제가능
	else
		if not(C_ADMIN_AUTH) then
			sql = sql & " and currstate<3"
		end if
	end if

	'response.write sql &"<br>"
	dbget.execute sql

	sql = ""
	sql = "select top 1 masteridx , detailidx , orderno" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_detail" & vbcrlf
	sql = sql & " where cancelyn='N'"
	sql = sql & " and orderno = '"& orderno &"'" & vbcrlf

	'response.write sql &"<br>"
	'response.end
	rsget.open sql ,dbget ,1

	if not(rsget.eof) then
		masteridxtmp = false
		orderno = rsget("orderno")
	else
		masteridxtmp = true
	end if

	rsget.close()

	'//디테일이 전부 취소 라면 마스터도 취소 시킨다
	if masteridxtmp then
		sql = ""
		sql = "update db_shop.dbo.tbl_shopbeasong_order_master set" & vbcrlf
		sql = sql & " cancelyn='Y'" & vbcrlf
		sql = sql & " where orderno = '"& orderno &"'"

		'response.write sql &"<br>"
		dbget.execute sql
	end if

	' 마스터 일부출고 설정
	sql = "update m set" & vbcrlf
	sql = sql & " m.ipkumdiv = '7'" & vbcrlf
	sql = sql & " , m.beadaldate = getdate()" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
	sql = sql & " join (" & vbcrlf
	sql = sql & " 	select dd.masteridx" & vbcrlf
	sql = sql & " 	from db_shop.dbo.tbl_shopbeasong_order_detail dd" & vbcrlf
	sql = sql & " 	where dd.cancelyn='N'" & vbcrlf
	sql = sql & " 	and dd.currstate < 7" & vbcrlf		' 출고완료 이전인거
	sql = sql & " 	and dd.orderno = '"& trim(orderno) &"'" & vbcrlf
	sql = sql & " 	group by dd.masteridx" & vbcrlf
	sql = sql & " ) as t" & vbcrlf
	sql = sql & " 	on m.masteridx = t.masteridx" & vbcrlf
	sql = sql & " where m.cancelyn='N'" & vbcrlf
	sql = sql & " and m.orderno = '"& trim(orderno) &"'" & vbcrlf

	'response.write sql &"<br>"
    dbget.Execute sql

	' 마스터 출고완료 설정
	sql = "update m set" & vbcrlf
	sql = sql & " m.ipkumdiv = '8'" & vbcrlf
	sql = sql & " , m.beadaldate = getdate()" & vbcrlf
	sql = sql & " from db_shop.dbo.tbl_shopbeasong_order_master m" & vbcrlf
	sql = sql & " left join (" & vbcrlf
	sql = sql & " 	select dd.masteridx" & vbcrlf
	sql = sql & " 	from db_shop.dbo.tbl_shopbeasong_order_detail dd" & vbcrlf
	sql = sql & " 	where dd.cancelyn='N'" & vbcrlf
	sql = sql & " 	and dd.currstate < 7" & vbcrlf		' 출고완료 이전인거
	sql = sql & " 	and dd.orderno = '"& trim(orderno) &"'" & vbcrlf
	sql = sql & " 	group by dd.masteridx" & vbcrlf
	sql = sql & " ) as t" & vbcrlf
	sql = sql & " 	on m.masteridx = t.masteridx" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " where m.cancelyn='N'" & vbcrlf
	sql = sql & " and m.orderno = '"& trim(orderno) &"'" & vbcrlf
	sql = sql & " and t.masteridx is null"

	'response.write sql &"<br>"
    dbget.Execute sql

	If Err.Number = 0 Then
	    dbget.CommitTrans

		'//마스터가 취소 됐으므로,출고 리스트  페이지로 넘긴다.
		if masteridxtmp then
			response.write "<script type='text/javascript'>"
		response.write "	alert('처리되었습니다');"
			response.write "	location.href='/common/offshop/beasong/shopbeasong_list.asp?orderno="&orderno&"&menupos="& menupos &"';"
			response.write "</script>"
			dbget.close()	:	response.End
		else
			response.write "<script type='text/javascript'>"
			response.write "	alert('처리되었습니다');"
			response.write "	location.href='/common/offshop/beasong/shopbeasong_input.asp?orderno="&orderno&"&menupos="& menupos &"';"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

	else
	    dbget.rollbackTrans

		response.write "<script type='text/javascript'>"
		response.write "	alert('값이 일치 하지 않습니다. 관리자 문의 하세요');"
		response.write "	history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

elseif (mode="SongjangInput") or (mode="SongjangInputCSV") then
	dim referer
	referer = request.ServerVariables("HTTP_REFERER")

	ordernoArr = request.Form("ordernoArr")
	songjangnoArr  = request.Form("songjangnoArr")
	songjangdivArr = request.Form("songjangdivArr")
	detailidxArr   = request.Form("detailidxArr")
	detailidx      = request.Form("detailidx")

	if (mode="SongjangInputCSV") then
	    ''CSV 입력은 끝에 , 가 하나 없음. 콤마 사이에 공백 있음
	    ordernoArr = Replace(ordernoArr," ","") & ","
	    songjangnoArr  = Replace(songjangnoArr," ","") & ","
	    songjangdivArr = Replace(songjangdivArr," ","") & ","
	    detailidxArr   = Replace(detailidxArr," ","") & ","
	end if

	TotAssignedRow = 0
	AssignedRow    = 0
	FailRow        = 0

    RectdetailidxArr   = split(detailidxArr,",")
    RectordernoArr = split(ordernoArr,",")
    RectSongjangnoArr  = split(songjangnoArr,",")
    RectSongjangdivArr = split(songjangdivArr,",")

    if IsArray(RectdetailidxArr) then
        OrderCount = Ubound(RectdetailidxArr)

        if (OrderCount<>Ubound(RectordernoArr)) or (OrderCount<>Ubound(RectSongjangnoArr)) or (OrderCount<>Ubound(RectSongjangdivArr)) then
            response.write "<script>alert('전송된 데이터가 일치하지 않습니다.');</script>"
            dbget.close()	:	response.end
        end if

    end if

    if Right(detailidxArr,1)="," then detailidxArr = Left(detailidxArr,Len(detailidxArr)-1)
    if (Right(ordernoArr,1)=",") then ordernoArr=Left(ordernoArr,Len(ordernoArr)-1)
    ordernoArr = replace(ordernoArr,",","','")

    dim tmp
    dbget.beginTrans

    ''송장번호입력 루프
    for i=0 to OrderCount - 1
        if (Trim(RectdetailidxArr(i))<>"") then

            ''품절출고 불가 등록된경우 SKIP
            mibeasongSoldOutExists = false

            'sqlStr = "select count(*) as CNT" & VbCRLF
            'sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_mibeasong_list" & VbCRLF
            'sqlStr = sqlStr + " where detailidx=" & Trim(RectdetailidxArr(i))  & VbCRLF
            'sqlStr = sqlStr + " and orderno='" & Trim(RectordernoArr(i)) & "'" & VbCRLF
            'sqlStr = sqlStr + " and code='05'" & VbCRLF

            'response.write sqlStr &"<br>"
            'rsget.CursorLocation = adUseClient
            'rsget.Open sqlStr, dbget, adOpenForwardOnly

        	'if Not rsget.Eof then
            '    mibeasongSoldOutExists = rsget("CNT")>0
            'end if

        	'rsget.close

        	if (mibeasongSoldOutExists) then
        	    FailRow = FailRow + 1
        	ELSE

                ''중복메일 방지용.
                sqlStr = ""
                sqlStr = "select d.masteridx"
                sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_detail d"
                sqlStr = sqlStr + " Join db_shop.dbo.tbl_shopbeasong_order_master m"
                sqlStr = sqlStr + " on d.masteridx=m.masteridx"
                sqlStr = sqlStr + " where d.orderno='" & Trim(RectordernoArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " and d.detailidx =" & Trim(RectdetailidxArr(i))  & VbCRLF
            	sqlStr = sqlStr + " and m.shopid='" & loginidshopormaker & "'"
            	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
            	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.

            	if (mode="SongjangInputCSV") then
            	    sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"
                end if

            	'response.write sqlStr &"<br>"
            	rsget.CursorLocation = adUseClient
                rsget.Open sqlStr, dbget, adOpenForwardOnly

            	if Not rsget.Eof then
            		tmp = ""
            		tmp = rsget("masteridx")&","

            	    if Not (InStr(iMailmasteridxArr,tmp)>0) then
            	        iMailmasteridxArr = iMailmasteridxArr + tmp
            	    end if
            	    tmp = ""
            	end if

            	rsget.close

                sqlStr = ""
            	sqlStr = "update D" & VbCRLF
            	sqlStr = sqlStr + " set currstate='7'" & VbCRLF
            	sqlStr = sqlStr + " ,songjangno='" & Trim(RectSongjangnoArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " ,songjangdiv='" & Trim(RectSongjangdivArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" & VbCRLF
            	sqlStr = sqlStr + " ,passday=IsNULL(db_sitemaster.dbo.fn_Ten_NetWorkDays(("
            	sqlStr = sqlStr + " 	select convert(varchar(10),baljudate,21)"
				sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shopbeasong_order_master mm"
            	sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail dd"
            	sqlStr = sqlStr + " 	on mm.masteridx = dd.masteridx"
            	sqlStr = sqlStr + "		where dd.detailidx=" & Trim(RectdetailidxArr(i)) & ""
            	sqlStr = sqlStr + " 	),IsNULL(convert(varchar(10),d.beasongdate,21),convert(varchar(10),getdate(),21))),0)"& VbCRLF
                sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_detail d"& VbCRLF
            	sqlStr = sqlStr + " Join db_shop.dbo.tbl_shopbeasong_order_master m"
                sqlStr = sqlStr + " on m.masteridx=d.masteridx"
            	sqlStr = sqlStr + " where d.orderno='" & Trim(RectordernoArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " and d.detailidx =" & Trim(RectdetailidxArr(i))  & VbCRLF
            	sqlStr = sqlStr + " and m.shopid='" & loginidshopormaker & "'"
            	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
            	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.

            	if (mode="SongjangInputCSV") then
            	    sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"   ''완료후 송장번호 변경 할 수 있음.. :: 개별입력만 가능하도록.
                end if

				'response.write sqlStr &"<br>"
                dbget.Execute sqlStr, AssignedRow

                TotAssignedRow = TotAssignedRow + AssignedRow

                if (AssignedRow=0) then FailRow = FailRow + 1
            END IF
        end if

    next

	'마스터 일부출고 설정
    sqlStr = " update 																					" & VbCRLF
    sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopbeasong_order_master 									" & VbCRLF
    sqlStr = sqlStr + " set 																			" & VbCRLF
    sqlStr = sqlStr + " 	ipkumdiv='7' 																" & VbCRLF
    sqlStr = sqlStr + " 	, beadaldate=getdate() 														" & VbCRLF
    sqlStr = sqlStr + " where 																			" & VbCRLF
    sqlStr = sqlStr + " 	masteridx in ( 																" & VbCRLF
    sqlStr = sqlStr + " 		select 																	" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		from 																	" & VbCRLF
    sqlStr = sqlStr + " 			db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
    sqlStr = sqlStr + " 			join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
    sqlStr = sqlStr + " 			on 																	" & VbCRLF
    sqlStr = sqlStr + " 				m.masteridx=d.masteridx 										" & VbCRLF
    sqlStr = sqlStr + " 		where 																	" & VbCRLF
    sqlStr = sqlStr + " 			1 = 1 																" & VbCRLF
    sqlStr = sqlStr + " 			and d.itemid<>0 													" & VbCRLF
    sqlStr = sqlStr + " 			and m.masteridx in ( 												" & VbCRLF
    sqlStr = sqlStr + " 				select distinct 												" & VbCRLF
    sqlStr = sqlStr + " 					m.masteridx 												" & VbCRLF
    sqlStr = sqlStr + " 				from 															" & VbCRLF
    sqlStr = sqlStr + " 					db_shop.dbo.tbl_shopbeasong_order_master m 					" & VbCRLF
    sqlStr = sqlStr + " 					join db_shop.dbo.tbl_shopbeasong_order_detail d 			" & VbCRLF
    sqlStr = sqlStr + " 					on 															" & VbCRLF
    sqlStr = sqlStr + " 						m.masteridx=d.masteridx 								" & VbCRLF
    sqlStr = sqlStr + " 				where 															" & VbCRLF
    sqlStr = sqlStr + " 					1 = 1 														" & VbCRLF
    sqlStr = sqlStr + " 					and d.detailidx in (" & detailidxArr & ") 					" & VbCRLF
    sqlStr = sqlStr + " 					and m.cancelyn='N' 											" & VbCRLF
    sqlStr = sqlStr + " 					and d.itemid<>0 											" & VbCRLF
    sqlStr = sqlStr + " 			) 																	" & VbCRLF
    sqlStr = sqlStr + " 		group by 																" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )>0 " & VbCRLF
    sqlStr = sqlStr + " 	) 																			" & VbCRLF

    'response.write sqlStr &"<br>"
    dbget.Execute sqlStr

	'전부출고
    sqlStr = " update 																					" & VbCRLF
    sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopbeasong_order_master 									" & VbCRLF
    sqlStr = sqlStr + " set 																			" & VbCRLF
    sqlStr = sqlStr + " 	ipkumdiv='8' 																" & VbCRLF
    sqlStr = sqlStr + " 	, beadaldate=getdate() 														" & VbCRLF
	sqlStr = sqlStr + " where 																			" & VbCRLF
    sqlStr = sqlStr + " 	masteridx in ( 																" & VbCRLF
    sqlStr = sqlStr + " 		select 																	" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		from 																	" & VbCRLF
    sqlStr = sqlStr + " 			db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
    sqlStr = sqlStr + " 			join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
    sqlStr = sqlStr + " 			on 																	" & VbCRLF
    sqlStr = sqlStr + " 				m.masteridx=d.masteridx 										" & VbCRLF
    sqlStr = sqlStr + " 		where 																	" & VbCRLF
    sqlStr = sqlStr + " 			1 = 1 																" & VbCRLF
    sqlStr = sqlStr + " 			and d.itemid<>0 													" & VbCRLF
    sqlStr = sqlStr + " 			and m.masteridx in ( 												" & VbCRLF
    sqlStr = sqlStr + " 				select distinct 												" & VbCRLF
    sqlStr = sqlStr + " 					m.masteridx 												" & VbCRLF
    sqlStr = sqlStr + " 				from 															" & VbCRLF
    sqlStr = sqlStr + " 					db_shop.dbo.tbl_shopbeasong_order_master m 					" & VbCRLF
    sqlStr = sqlStr + " 					join db_shop.dbo.tbl_shopbeasong_order_detail d 			" & VbCRLF
    sqlStr = sqlStr + " 					on 															" & VbCRLF
    sqlStr = sqlStr + " 						m.masteridx=d.masteridx 								" & VbCRLF
    sqlStr = sqlStr + " 				where 															" & VbCRLF
    sqlStr = sqlStr + " 					1 = 1 														" & VbCRLF
    sqlStr = sqlStr + " 					and d.detailidx in (" & detailidxArr & ") 					" & VbCRLF
    sqlStr = sqlStr + " 					and m.cancelyn='N' 											" & VbCRLF
    sqlStr = sqlStr + " 					and d.itemid<>0 											" & VbCRLF
    sqlStr = sqlStr + " 			) 																	" & VbCRLF
    sqlStr = sqlStr + " 		group by 																" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0 " & VbCRLF
    sqlStr = sqlStr + " 	) 																			" & VbCRLF

    'response.write sqlStr &"<br>"
    dbget.Execute sqlStr

    ''메일보내기 루프
    iMailmasteridxArr = split(iMailmasteridxArr,",")

    if IsArray(iMailmasteridxArr) then
        for i=LBound(iMailmasteridxArr) to UBound(iMailmasteridxArr)

            if Trim(iMailmasteridxArr(i))<>"" then
                if (application("Svr_Info")<>"Dev") then
                    'call fcSendMailFinish_Dlv_Designer_off(iMailmasteridxArr(i),MakerID)
                end if
            end if
        next
    end if



	'문자발송
	dim buyhparr
	songjangdivarr = ""
	songjangnoarr = ""

    sqlStr = " select distinct 															" & VbCRLF
    sqlStr = sqlStr + " 	m.masteridx 												" & VbCRLF
    sqlStr = sqlStr + " 	, m.buyhp 													" & VbCRLF
    sqlStr = sqlStr + " 	, d.songjangdiv 											" & VbCRLF
    sqlStr = sqlStr + " 	, d.songjangno 												" & VbCRLF
    sqlStr = sqlStr + " from 															" & VbCRLF
    sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopbeasong_order_master m 					" & VbCRLF
    sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail d 			" & VbCRLF
    sqlStr = sqlStr + " 	on 															" & VbCRLF
    sqlStr = sqlStr + " 		m.masteridx=d.masteridx 								" & VbCRLF
    sqlStr = sqlStr + " where 															" & VbCRLF
    sqlStr = sqlStr + " 	1 = 1 														" & VbCRLF
    sqlStr = sqlStr + " 	and d.detailidx in (" & detailidxArr & ") 					" & VbCRLF
    sqlStr = sqlStr + " 	and m.cancelyn='N' 											" & VbCRLF
    sqlStr = sqlStr + " 	and d.itemid<>0 											" & VbCRLF

	rsget.open sqlStr ,dbget ,1

	if not(rsget.eof) then
		do until rsget.Eof
			buyhparr 		= buyhparr + "," + rsget("buyhp")
			songjangdivarr 	= songjangdivarr + "," + CStr(rsget("songjangdiv"))
			songjangnoarr	= songjangnoarr + "," + CStr(rsget("songjangno"))
			rsget.MoveNext
		loop
	end if
	rsget.close()

    buyhparr = split(buyhparr,",")
    songjangdivarr = split(songjangdivarr,",")
    songjangnoarr = split(songjangnoarr,",")

    if IsArray(buyhparr) then
        for i=LBound(buyhparr) to UBound(buyhparr)
            if Trim(buyhparr(i))<>"" then
                if (application("Svr_Info")<>"Dev") then
                    'call SendNormalSMS(Trim(buyhparr(i)), "", "[텐바이텐샵] 상품이 출고되었습니다. [" & DeliverDivCd2Nm(Trim(songjangdivarr(i))) & "]" & Trim(songjangnoarr(i)) & "")
                    Call SendNormalSMS_LINK(Trim(buyhparr(i)), "1644-6030", "[텐바이텐샵] 상품이 출고되었습니다. [" & DeliverDivCd2Nm(Trim(songjangdivarr(i))) & "]" & Trim(songjangnoarr(i)) & "")
                end if
            end if
        next
    end if

	If Err.Number = 0 Then
	    dbget.CommitTrans
	Else
	    dbget.RollBackTrans
	End If

    dim AlertMsg
    AlertMsg = TotAssignedRow & "건 처리 되었습니다."
    if (FailRow>0) then
        AlertMsg = AlertMsg & "\n\n(" & FailRow & "건 입력 실패)"
    end if

    response.write "<script type='text/javascript'>alert('" & AlertMsg & "')</script>"

    if (mode="SongjangInputCSV") then
        response.write "<script type='text/javascript'>opener.location.reload();</script>"
        response.write "<script type='text/javascript'>window.close();</script>"
    else
        response.write "<script type='text/javascript'>location.replace('" + CStr(referer) + "')</script>"
    end if
    dbget.close()	:	response.End

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('잘못된 경로를 지정 하셨습니다.');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close()	:	response.End
end if
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
' 차후 사용
'    sql = "select max(isnull(baljunum,0)) as maxbaljunum, convert(varchar,getdate(),109) as baljudate" & vbcrlf
'    sql = sql & " from [db_storage].[dbo].tbl_shopbalju_customer"
'
'    'response.write sql & "<Br>"
'	rsget.Open sqlStr, dbget, 1
'	if Not rsget.Eof then
'		baljunum = rsget("maxbaljunum") + 1
'		baljudate = rsget("baljudate")
'	end if
'	rsget.close
'
'	sql = "select (IsNull(max(differencekey), 0) + 1) as differencekey" & vbcrlf
'	sql = sql & " from [db_storage].[dbo].tbl_shopbalju_customer" & vbcrlf
'	sql = sql & " where convert(varchar(10),baljudate,21)=convert(varchar(10),getdate(),21)"
'
'    'response.write sql & "<Br>"
'	rsget.Open sqlStr,dbget,1
'		differencekey = rsget("differencekey")
'	rsget.close
'
'	ordercnt = ubound(ordernoarr)
'
'	for i = 0 to ordercnt
'        sql = "insert into [db_storage].[dbo].tbl_shopbalju_customer(baljunum, baljuid, orderno, baljudate, differencekey, workgroup, songjangdiv)" & vbcrlf
'        sql = sql & " values("& baljunum &", '"& trim(shopidarr(i)) &"', '"& trim(ordernoarr(i)) &"', convert(datetime,'" + CStr(baljudate) + "',109), " + CStr(differencekey) + ", '" + CStr(workgroup) + "', " + CStr(songjangdiv) + ") "
'
'		'response.write sql &"<br>"
'		dbget.execute sql
'	next
'
'	'response.write sql &"<br>"
'	dbget.execute sql
%>