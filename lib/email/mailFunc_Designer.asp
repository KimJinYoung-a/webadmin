<!-- #include virtual="/lib/email/mailFunction.asp" -->
<%
'+--------------------------------------------------------------------------------------------------------------------------------+
'|                                        업체 배송 상품 메일 발송                                                                |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'|             함 수 명                               |                          기    능                                         |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer(orderserial,makerid) | 출고 메일 발송(업체배송 출고)                                             |
'|                                                    | 사용예 : fcSendMailFinish_Dlv_Designer('012012304','1293495006')          |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer_off(detailidx,makerid)  오프라인 출고 메일 발송(업체배송 출고)                                   |
'|                                                    | 사용예 : fcSendMailFinish_Dlv_Designer_off('012012304','1293495006')      |
'+----------------------------------------------------+---------------------------------------------------------------------------+

'' 해당브랜드 전체 출고시 발송(중복 발송 제거) ''//2014/03/31 추가 , 2019/06/27 최근 1시간내 발송건이 있는경우만.
function isDlvFinishedByBrand(vOrderSerial,vMakerid)
    dim strSQL, targetCNT, DLVCNT, recentDLVCNT
    targetCNT = 0
    DLVCNT    = 0
	recentDLVCNT =0

    strSQL = " select count(*) as targetCNT"
    strSQL = strSQL & " , sum(CASE WHEN d.currstate=7 and beasongdate is Not NULL THEN 1 ELSE 0 END) as DLVCNT" &VbCRLF
	strSQL = strSQL & " , sum(CASE WHEN d.beasongdate>dateadd(n,-60,getdate()) THEN 1 ELSE 0 END) as recentDLVCNT" &VbCRLF
    strSQL = strSQL & " from [db_order].[dbo].tbl_order_master m" &VbCRLF
    strSQL = strSQL & " 	Join [db_order].[dbo].tbl_order_detail d" &VbCRLF
    strSQL = strSQL & " 	on m.orderserial=d.orderserial" &VbCRLF
    strSQL = strSQL & " where d.itemid not in (0,100)" &VbCRLF
    strSQL = strSQL & " and d.orderserial='"&vOrderSerial&"'" &VbCRLF
    strSQL = strSQL & " and d.makerid='"&vMakerid&"'" &VbCRLF
    strSQL = strSQL & " and d.cancelyn<>'Y'" &VbCRLF
    strSQL = strSQL & " and m.cancelyn='N'" &VbCRLF
    rsget.CursorLocation = adUseClient
    rsget.Open strSQL, dbget, adOpenForwardOnly
    IF  not rsget.Eof  THEN
        targetCNT = rsget("targetCNT")
        DLVCNT    = rsget("DLVCNT")
		recentDLVCNT = rsget("recentDLVCNT")
    END IF
    rsget.CLOSE

    isDlvFinishedByBrand = false
    if (DLVCNT<1) or (targetCNT<>DLVCNT) or (recentDLVCNT<1) then EXIT Function

    isDlvFinishedByBrand = true
end function

Function fcSendMailFinish_Dlv_Designer(vOrderSerial,vMakerid)	'/2011.04.21 한용민 수정

	IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub ,strHTML_MAINother
	' 배송 주체별 HTML
	strHTML_MAIN = ""
	strHTML_MAIN = strHTML_MAIN &"<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"	<td style=""padding:0 29px 45px; margin:0;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"			<th style=""margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, '돋움', sans-serif; text-align:left; color:#000;"">발송된 상품 <span style=""margin-left:15px; padding:0; font-size:11px; line-height:11px; font-weight:normal; font-family:dotum, '돋움', sans-serif; vertical-align:2px; color:#808080; text-align:left;"">운송장 번호를 클릭하시면 배송현황을 확인하실 수 있습니다.</span><th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		</tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"			<td style=""border-top:solid 2px #000;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"				<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; font-size:12px; font-family:dotum, '돋움', sans-serif; color:#707070;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"					<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, '돋움', sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;"">상품</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">상품코드</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:250px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">상품명[옵션]</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:37px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">수량</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:95px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">주문상태</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:108px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">택배정보</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"					</tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"				[$ITEMHTMLTABLE$]"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"				</table>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"			</td>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		</tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		</table>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"	</td>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"</tr>"

	strHTML_MAINother = ""
	strHTML_MAINother = strHTML_MAINother & "<tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "	<td style=""padding:0 29px 45px; margin:0;"">"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		<tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "			<th style=""margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, '돋움', sans-serif; text-align:left; color:#000;"">함께 주문하신 상품 배송현황<th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		</tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		<tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "			<td style=""border-top:solid 2px #000;"">"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "				<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; font-size:12px; font-family:dotum, '돋움', sans-serif; color:#707070;"">"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "					<tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, '돋움', sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;"">상품</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">상품코드</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:250px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">상품명[옵션]</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:37px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">수량</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:95px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">주문상태</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:108px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px; line-height:12px;"">택배정보</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "					</tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "				[$ITEMHTMLTABLE$]"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "				</table>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "			</td>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		</tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother &"		</table>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "	</td>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "</tr>"

	' 기본 상품 설명부분 HTML
	strHTML_Sub =""
	strHTML_Sub = strHTML_Sub & "<tr>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:50px; padding:6px 0;border-bottom:solid 1px #eaeaea;"">"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		<img src=""[$ITEM_IMAGE_URL$]"" width=50 height=50 alt="""" />"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:100px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; color:#707070; font-size:11px; line-height:11px; font-family:dotum, '돋움', sans-serif;"">"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		[$ITEM_ID$]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:250px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; color:#707070; font-size:11px; line-height:17px; font-family:dotum, '돋움', sans-serif;"">"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		[[$ITEM_brandName$]]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		<br /> [$ITEM_NAME$]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:37px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:13px; line-height:13px;"">[$ITEM_QUANTITY$]</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	[$ITEM_DLV_STATUS$]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:108px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:12px; text-align:center; font-family:dotum, '돋움', sans-serif; text-align:center;"">[$ITEM_DELIVERY_LINK$]</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "</tr>"

    '주문 상품 정보
	dim strSQL
	dim ITIMG , ITNM , ITID , ITOPNM , ITNO , ITbrandName ,ITmakerid
	dim DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '지금 배송,같이주문한 상품

	tmpHTML="":NowHTML="":OtherHTML=""

	strSQL =" SELECT a.itemid, a.itemoptionname, c.smallimage, c.itemname,c.makerid ," &_
			" (c.cate_large + c.cate_mid + c.cate_small) as itemserial," &_
			" a.itemcost as sellcash, a.itemno, a.isupchebeasong, a.songjangdiv, replace(isnull(a.songjangno,''),'-','') as songjangno, a.currstate" &_
			" ,s.divname,s.findurl ,c.brandName" &_
			" FROM [db_order].[dbo].tbl_order_detail a" &_
			" JOIN [db_item].[dbo].tbl_item c" &_
			" 	on c.itemid = a.itemid" &_
			" LEFT JOIN db_order.[dbo].tbl_songjang_div s" &_
			" 	on a.songjangdiv=s.divcd" &_
			" WHERE a.orderserial = '" & vOrderSerial & "'" &_
			" and a.itemid <> '0'" &_
			" and (a.cancelyn<>'Y')"

	'response.write strSQL
	rsget.Open strSQL,dbget,1

	IF  not rsget.Eof  THEN
		rsget.Movefirst

		DO UNTIL rsget.eof

			'-- 브랜드
			ITmakerid = db2html(rsget("makerid"))

			'-- 브랜드명
			ITbrandName = db2html(rsget("brandName"))

			'--- 상품이미지
			ITIMG = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
			' 상품 코드
			ITID = rsget("itemid")
			'--- 상품명
			ITNM = db2html(rsget("itemname"))
			'--- 상품옵션명
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & " [" & ITOPNM & "]"
			END IF
			'--- 상품수량 -- 수량별 style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = Cstr(rsget("itemno"))
			END IF

			'--- 배송상태 지정
				IF rsget("currstate") = 7 THEN
					DLVSTS = "<td style=""width:95px; height:44px; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '돋움', sans-serif; color:#dd5555; font-size:12px;"">출고완료</td>"
				 ELSE
					DLVSTS = "<td style=""width:95px; height:44px; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '돋움', sans-serif; color:#707070; font-size:12px;"">상품준비중</td>"
				 END IF
			'--- 택배/송장 설정
			IF ((Not isnull(rsget("songjangno"))) and  (rsget("songjangno")<>"") ) THEN
				DLVLKTXT = ""
				DLVLKTXT = DLVLKTXT & "<span style=""margin:0; padding:0; color:#707070; font-size:12px; font-weight:bold; line-height:18px; font-family:dotum, '돋움', sans-serif; text-align:center;"">" & db2html(rsget("divname")) & "</span><br />"
				DLVLKTXT = DLVLKTXT & "<a href=""" & db2html(rsget("findurl")) & rsget("songjangno") & """ style=""margin:0; padding:0; font-size:12px; color:#dd5555; font-size:11px; line-height:18px; font-family:dotum, '돋움', sans-serif; color:#0066cc; text-align:center;"">" & rsget("songjangno") & "</a>"
			else
				DLVLKTXT ="-"
			end if
			tmpHTML = strHTML_Sub
			tmpHTML = replace(tmpHTML,"[$ITEM_makerid$]",ITmakerid)
			tmpHTML = replace(tmpHTML,"[$ITEM_brandName$]",ITbrandName)
			tmpHTML = replace(tmpHTML,"[$ITEM_IMAGE_URL$]",ITIMG)
			tmpHTML = replace(tmpHTML,"[$ITEM_ID$]",ITID)
			tmpHTML = replace(tmpHTML,"[$ITEM_NAME$]",ITNM)
			tmpHTML = replace(tmpHTML,"[$ITEM_QUANTITY$]",ITNO)
			tmpHTML = replace(tmpHTML,"[$ITEM_DLV_STATUS$]",DLVSTS)
			tmpHTML = replace(tmpHTML,"[$ITEM_DELIVERY_LINK$]",DLVLKTXT)

			IF rsget("isupchebeasong") = "Y" and rsget("makerid")=vMakerid and rsget("songjangno")<>"" THEN
				NowHTML= NowHTML & tmpHTML
				isNowDLV= true
			ELSE
				OtherHTML = OtherHTML & tmpHTML
				isOtherDLV= true
			END IF

			tmpHTML ="":ITIMG="":ITID="":ITNM="":ITOPNM="":ITNO="":DLVSTS="":DLVLKTXT=""

			rsget.movenext
		LOOP
    ELSE
    	rsget.close
		EXIT FUNCTION

    END IF
    rsget.close

	IF NowHTML<>"" and isNowDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_shiped.gif"" alt=""출고된 상품의 배송정보"">"
		NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
		NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		NowHTML= ""
	END IF

	IF OtherHTML<>"" and isOtherDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_otherpd.gif"" alt="" 같이 주문하신 상품 배송현황"">"
		OtherHTML = replace(strHTML_MAINother,"[$ITEMHTMLTABLE$]",OtherHTML)
		OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		OtherHTML=""
	END IF


	'//=======  메일정보 & 배송정보 , 결제정보 불러오기 =========/
	'// ( !!!!! /lib/email/mailFunction.asp 참조 !!!!! )
	call getInfo(vOrderSerial)

	IF MailTo ="" Then
		Exit Function
	End IF

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	oMail.MailType		 = 8 '메일 종류별 고정값 (mailLib2.asp 참고)
	oMail.MailTitles	 = "[텐바이텐]주문하신 상품에 대한 텐바이텐 배송안내입니다!"
	'oMail.SenderNm		 = "텐바이텐"
	'oMail.SenderMail	 = "customer@10x10.co.kr"
	oMail.AddrType		 = "string"
	oMail.ReceiverNm	 = MailTo_Nm
	oMail.ReceiverMail	 = MailTo

	MailHTML= oMail.getMailTemplate()

	IF MailHTML="" Then
		SET oMail = nothing
		response.write "<script>alert('메일발송이 실패 하였습니다.');</script>"
		Exit Function
    End IF

	'// 실제 메일에 정보 치환
	MailHTML = replace(MailHTML,"[$USER_NAME$]", MailTo_Nm) ' 주문자 이름
	MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' 주문번호
	MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '출고된 상품 HTML
	MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'같이 주문한상품 HTML
	MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",newReqInfoHTML)	'배송지 정보 HTML
	MailHTML = replace(MailHTML,"http://mailzine.10x10.co.kr/2017/txt_noti_send_prd.png", "http://mailzine.10x10.co.kr/2017/txt_noti_send_prd2.png")	' 업배일경우 배송 이미지 변경. 업체가 가라로 송장찍고 늦게 보내는 케이스가 있음.

	oMail.MailConts = MailHTML

	'response.write MailHTML
	'response.end
	oMail.MailerMailGubun = 4		' 메일러 자동메일 번호
	oMail.Send_TMSMailer()		'TMS메일러
	'oMail.Send_Mailer()
	''oMail.Send_CDO()
	'oMail.Send_CDONT()

	SET oMail = nothing

End Function

Function fcSendMailFinish_Dlv_Designer_off(vmasteridx,vMakerid)
	IF trim(vmasteridx) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub

	dim vOrderSerial

	' 배송 주체별 HTML
	strHTML_MAIN ="" &_
		"<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
		"<tr>" &_
		"	<td style=""padding-bottom:7px;"">[$DELIVERY_HOST_IMG$]</td>" &_
		"</tr>" &_
		"<tr>" &_
		"	<td>" &_
		"		<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"" style=""border-bottom:1px solid #dddddd"">" &_
		"		[$ITEMHTMLTABLE$]" &_
		"		</table>" &_
		"	</td>" &_
		"</tr>" &_
		"</table>"

	' 기본 상품 설명부분 HTML '이미지"<td><img src=""[$ITEM_IMAGE_URL$]"" width=""50"" height=""50""></td>" &_
	strHTML_Sub ="" &_
			"<tr>" &_
			"	<td>" &_
			"		<table width=""548"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top:1px solid #dddddd"">" &_
			"		<tr>" &_
			"			<td width=""260"" align=""right"" style=""border-right: 1px solid #dddddd"">" &_
			"				<table width=""255"" height=""50""  border=""0"" cellpadding=""0"" cellspacing=""0"">" &_
			"				<tr>" &_
			"					<td width=""50"" valign=""bottom"">" &_
			"						<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
			"						<tr>" &_

			"						</tr>" &_
			"						</table>" &_
			"					</td>" &_
			"					<td  style=""padding:5"">[$ITEM_ID$]<br>[$ITEM_NAME$] </td>" &_
			"				</tr>" &_
			"				</table>" &_
			"			</td>" &_
			"			<td align=""center"">" &_
			"				<table width=""100%"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#eeeeee"">" &_
			"				<tr>" &_
			"					<td width=""60"" height=""35"" align=""center"">수 량</td>" &_
			"					<td width=""60"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF"">[$ITEM_QUANTITY$]</td>" &_
			"					<td width=""60"" align=""center"" style=""padding:0 5 0 5;"">배송현황</td>" &_
			"					<td class=""black12px"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF""> [$ITEM_DLV_STATUS$]</td>" &_
			"				</tr>" &_
			"				<tr height=""1"">" &_
			"					<td colspan=""4"" align=""center"" bgcolor=""#dddddd""></td>" &_
			"				</tr>" &_
			"				<tr>" &_
			"					<td align=""center"">운송장</td>" &_
			"					<td colspan=""3"" style=""padding:5"" bgcolor=""#FFFFFF""><strong class=""Information_font"">[$ITEM_DELIVERY_LINK$]</strong></td>" &_
			"				</tr>" &_
			"				</table>" &_
			"			</td>" &_
			"		</tr>" &_
			"		</table>" &_
			"	</td>" &_
			"</tr>"

    '주문 상품 정보
	dim strSQL, ITIMG , ITNM , ITID , ITOPNM , ITNO ,DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '지금 배송,같이주문한 상품

	tmpHTML="":NowHTML="":OtherHTML=""

	strSQL =" SELECT" &_
			" d.itemid, d.itemgubun,d.itemoption,d.makerid, d.itemno, d.isupchebeasong" &_
			" ,replace(isnull(d.songjangno,''),'-','') as songjangno, d.currstate, d.songjangdiv" &_
			" ,od.sellprice as sellcash,od.itemoptionname, od.itemname" &_
			" ,s.divname,s.findurl, m.orderno " &_
			" from db_shop.dbo.tbl_shopbeasong_order_master m" &_
			" join db_shop.dbo.tbl_shopbeasong_order_detail d" &_
			" on m.masteridx=d.masteridx" &_
			" left join [db_shop].[dbo].tbl_shopjumun_detail od" &_
			" on d.orgdetailidx = od.idx" &_
			" LEFT JOIN db_order.[dbo].tbl_songjang_div s" &_
			" 	on d.songjangdiv=s.divcd" &_
			" WHERE d.masteridx = " & vmasteridx & "" &_
			" and d.itemid not in (0,100)" &_
			" and (d.cancelyn<>'Y')"

	'response.write strSQL &"<br>"
	rsget.Open strSQL,dbget,1
	IF  not rsget.Eof  THEN
		rsget.Movefirst

		vOrderSerial = rsget("orderno")

		DO UNTIL rsget.eof

			'--- 상품이미지
			ITIMG = ""
			' 상품 코드
			ITID = rsget("itemgubun")&Format00(6,rsget("itemid"))&rsget("itemoption")
			'--- 상품명
			ITNM = db2html(rsget("itemname"))
			'--- 상품옵션명
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & "<br><font color=""blue"">[" & ITOPNM & "]</font>"
			END IF
			'--- 상품수량 -- 수량별 style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = "<strong>" & Cstr(rsget("itemno")) & "</strong>"
			END IF

			'--- 배송상태 지정
				IF rsget("currstate") = 7 THEN
					 DLVSTS = "<span class=""black12px"">출고완료</span>"
				 ELSE
					 DLVSTS = "상품준비중"
				 END IF
			'--- 택배/송장 설정
			IF ((Not isnull(rsget("songjangno"))) and  (rsget("songjangno")<>"") ) THEN
				DLVLKTXT ="<a href=""" & db2html(rsget("findurl")) & rsget("songjangno") & """ target=""_blank""  class=""link_title"">" & db2html(rsget("divname")) & " " & rsget("songjangno") & "</a>"
			else
				DLVLKTXT ="-"
			end if
			tmpHTML = strHTML_Sub
			'tmpHTML = replace(tmpHTML,"[$ITEM_IMAGE_URL$]",ITIMG)
			tmpHTML = replace(tmpHTML,"[$ITEM_ID$]",ITID)
			tmpHTML = replace(tmpHTML,"[$ITEM_NAME$]",ITNM)
			tmpHTML = replace(tmpHTML,"[$ITEM_QUANTITY$]",ITNO)
			tmpHTML = replace(tmpHTML,"[$ITEM_DLV_STATUS$]",DLVSTS)
			tmpHTML = replace(tmpHTML,"[$ITEM_DELIVERY_LINK$]",DLVLKTXT)

			IF rsget("isupchebeasong") = "Y" and rsget("makerid")=vMakerid and rsget("songjangno")<>"" THEN
				NowHTML= NowHTML & tmpHTML
				isNowDLV= true
			ELSE
				OtherHTML = OtherHTML & tmpHTML
				isOtherDLV= true
			END IF

			tmpHTML ="":ITIMG="":ITID="":ITNM="":ITOPNM="":ITNO="":DLVSTS="":DLVLKTXT=""

			rsget.movenext
		LOOP
    ELSE

    	rsget.close
		EXIT FUNCTION

    END IF
    rsget.close

	IF NowHTML<>"" and isNowDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text01.gif"" width=""79"" height=""18"" alt=""출고된 상품의 배송정보"">"
		NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
		NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		NowHTML= ""
	END IF

	IF OtherHTML<>"" and isOtherDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text02.gif"" width=""193"" height=""18"" alt="" 같이 주문하신 상품 배송현황"">"
		OtherHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",OtherHTML)
		OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		OtherHTML=""
	END IF

	'//=======  메일정보 & 배송정보 , 결제정보 불러오기 =========/
	call getInfo_off(vmasteridx)

	IF MailTo ="" Then
		Exit Function
	End IF

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	oMail.MailType		 = 8 '메일 종류별 고정값 (mailLib2.asp 참고)
	oMail.MailTitles	 = "[텐바이텐샵]주문하신 상품에 대한 텐바이텐 배송안내입니다!"
	'oMail.SenderNm		 = "텐바이텐"
	'oMail.SenderMail	 = "customer@10x10.co.kr"
	oMail.AddrType		 = "string"
	oMail.ReceiverNm	 = MailTo_Nm
	oMail.ReceiverMail	 = MailTo

	MailHTML= oMail.getMailTemplate()

	IF MailHTML="" Then
		SET oMail = nothing
		response.write "<script>alert('메일발송이 실패 하였습니다.');</script>"
		Exit Function
    End IF

	'// 실제 메일에 정보 치환
	MailHTML = replace(MailHTML,"[$USER_NAME$]", MailTo_Nm) ' 주문자 이름
	MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' 주문번호
	MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '출고된 상품 HTML
	MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'같이 주문한상품 HTML
	MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'배송지 정보 HTML

	oMail.MailConts = MailHTML
	oMail.MailerMailGubun = 4		' 메일러 자동메일 번호
	oMail.Send_TMSMailer()		'TMS메일러
	'oMail.Send_Mailer()
	'oMail.Send_CDO()
	'oMail.Send_CDONT()

	SET oMail = nothing
End Function

%>
