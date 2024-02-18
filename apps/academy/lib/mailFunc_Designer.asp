
<%

'+----------------------------------------------------------------------------------------------------------------------+
'|                                        업체 배송 상품 메일 발송                                                      |
'+----------------------------------------------------+-----------------------------------------------------------------+
'|             함 수 명                               |                          기    능                               |
'+----------------------------------------------------+-----------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer(orderserial,makerid) | 출고 메일 발송(업체배송 출고)                                   |
'|                                                    | 사용예 : fcSendMailFinish_DlvTEN('012012304','1293495006')      |
'+----------------------------------------------------+-----------------------------------------------------------------+



Function fcSendMailFinish_Dlv_Designer(vOrderSerial,vMakerid)

		IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

		dim strHTML_MAIN,strHTML_Sub
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

		' 기본 상품 설명부분 HTML
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
				"							<td><img src=""[$ITEM_IMAGE_URL$]"" width=""50"" height=""50""></td>" &_
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
		dim strSQL
		dim ITIMG , ITNM , ITID , ITOPNM , ITNO
		dim DLVSTS, DLVLKTXT
		dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
		dim isNowDLV,isOtherDLV '지금 배송,같이주문한 상품

		tmpHTML="":NowHTML="":OtherHTML=""

		strSQL =" SELECT a.itemid, a.itemoptionname, c.smallimage, c.itemname,c.makerid ," &_
				" (c.cate_large + c.cate_mid + c.cate_small) as itemserial," &_
				" a.itemcost as sellcash, a.itemno, a.isupchebeasong, a.songjangdiv, replace(isnull(a.songjangno,''),'-','') as songjangno, a.currstate" &_
				" ,s.divname,s.findurl" &_
				" FROM [db_academy].[dbo].tbl_academy_order_detail a" &_
				" JOIN [db_academy].[dbo].tbl_diy_item c" &_
				" 	on c.itemid = a.itemid" &_
				" LEFT JOIN db_academy.[dbo].tbl_songjang_div s" &_
				" 	on a.songjangdiv=s.divcd" &_
				" WHERE a.orderserial = '" & vOrderSerial & "'" &_
				" and a.itemid <> '0'" &_
				" and (a.cancelyn<>'Y')"


		'response.write strSQL

		rsACADEMYget.Open strSQL,dbACADEMYget,1
		IF  not rsACADEMYget.Eof  THEN
			rsACADEMYget.Movefirst

			DO UNTIL rsACADEMYget.eof

				'--- 상품이미지
				ITIMG = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsACADEMYget("itemid")) & "/" & rsACADEMYget("smallimage")
				' 상품 코드
				ITID = rsACADEMYget("itemid")
				'--- 상품명
				ITNM = db2html(rsACADEMYget("itemname"))
				'--- 상품옵션명
				ITOPNM = db2html(rsACADEMYget("itemoptionname"))

				IF ITOPNM<>"" then
					ITNM = ITNM & "<br><font color=""blue"">[" & ITOPNM & "]</font>"
				END IF
				'--- 상품수량 -- 수량별 style
				ITNO = Cstr(rsACADEMYget("itemno"))
				IF rsACADEMYget("itemno")>1 THEN
					ITNO = "<strong>" & Cstr(rsACADEMYget("itemno")) & "</strong>"
				END IF

				'--- 배송상태 지정
					IF rsACADEMYget("currstate") = 7 THEN
						 DLVSTS = "<span class=""black12px"">출고완료</span>"
					 ELSE
						 DLVSTS = "상품준비중"
					 END IF
				'--- 택배/송장 설정
				IF ((Not isnull(rsACADEMYget("songjangno"))) and  (rsACADEMYget("songjangno")<>"") ) THEN
					DLVLKTXT ="<a href=""" & db2html(rsACADEMYget("findurl")) & rsACADEMYget("songjangno") & """ target=""_blank""  class=""link_title"">" & db2html(rsACADEMYget("divname")) & " " & rsACADEMYget("songjangno") & "</a>"
				else
					DLVLKTXT ="-"
				end if
				tmpHTML = strHTML_Sub
				tmpHTML = replace(tmpHTML,"[$ITEM_IMAGE_URL$]",ITIMG)
				tmpHTML = replace(tmpHTML,"[$ITEM_ID$]",ITID)
				tmpHTML = replace(tmpHTML,"[$ITEM_NAME$]",ITNM)
				tmpHTML = replace(tmpHTML,"[$ITEM_QUANTITY$]",ITNO)
				tmpHTML = replace(tmpHTML,"[$ITEM_DLV_STATUS$]",DLVSTS)
				tmpHTML = replace(tmpHTML,"[$ITEM_DELIVERY_LINK$]",DLVLKTXT)

				IF rsACADEMYget("isupchebeasong") = "Y" and rsACADEMYget("makerid")=vMakerid and rsACADEMYget("songjangno")<>"" THEN
					NowHTML= NowHTML & tmpHTML
					isNowDLV= true
				ELSE
					OtherHTML = OtherHTML & tmpHTML
					isOtherDLV= true
				END IF

				tmpHTML ="":ITIMG="":ITID="":ITNM="":ITOPNM="":ITNO="":DLVSTS="":DLVLKTXT=""

				rsACADEMYget.movenext
			LOOP
        ELSE
        	rsACADEMYget.close
			EXIT FUNCTION

        END IF
        rsACADEMYget.close

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
