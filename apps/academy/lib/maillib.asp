<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

'fcSendMail_OrderFinish("B0110377784")

'// 010-111-3333 => 010-***-3333
function AstarPhoneNumber(phoneNumber)
	Dim regEx, result
	if isNULL(phoneNumber) then Exit function
	    
	Set regEx = New RegExp

	With regEx
		.Pattern = "([0-9]+)-([0-9]+)-([0-9]+)"
		.IgnoreCase = True
		.Global = True
	End With

	result = regEx.Replace(phoneNumber,"$1-***-$3")

	if (result = phoneNumber) then
		if (Len(phoneNumber) >= 4) then
			result = Left(phoneNumber, (Len(phoneNumber) - 4)) & "****"
		end if
	end if

	set regEx = nothing

	AstarPhoneNumber = result
end Function

'// 홍길동 => 홍*동
function AstarUserName(userName)
	Dim result
    if isNULL(userName) then Exit function
        
	Select Case Len(userName)
		Case 0
			''
		Case 1
			result = "*"
		Case 2
			result = Left(userName,1) & "*"
		Case Else
			''3이상
			result = Left(userName,1) & "*" & Right(userName,1)
	End Select

	AstarUserName = result
end function

Public Function fcSendMail_OrderFinish(vOrderSerial)

		dim mailFrom, nameFrom, mailTitle, mailType

		mailFrom = "customer@thefingers.co.kr"
		nameFrom = "더핑거스"
		mailTitle = "[더핑거스] 주문이 완료되었습니다."
		mailType = "6"							'참고 mailLib2.asp

		Call fcSendMail(vOrderSerial, mailFrom, nameFrom, mailTitle, mailType)

End Function

Public Function fcSendMail_UpcheSendItem(vOrderSerial, vMakerid)

	'response.write vOrderSerial & "aaaaaaaaaaaaaa"
	Call fcSendMailFinish_Dlv_DIY(vOrderSerial,vMakerid)

End Function

Public Function fcSendMail_SendMiChulgoMail(detailidx)

		'response.write detailidx & "aaaaaaaaaaaaaa"
		Call SendMiChulgoMail(detailidx)

End Function










Public Function fcSendMail(vOrderSerial, mailFrom, nameFrom, mailTitle, mailType)

		'// 배송정보 & 메일정보

		dim strSQL

		dim buyerName, subTotalPrice, reqName , reqZipcode , reqAddress , reqPhone ,repHp , reqComment



		strSQL =" SELECT top 1 buyname,buyemail " &_
				" ,reqName,reqZipcode ,reqAddress ,reqPhone , reqhp , comment " &_
				" FROM [db_academy].[dbo].tbl_academy_order_master "  &_
				" WHERE orderserial = '" + vOrderSerial + "'"

		rsACADEMYget.Open strSQL,dbACADEMYget,1

		IF  not rsACADEMYget.EOF  THEN
			mailTo 		= rsACADEMYget("buyemail")
			buyerName  	= db2html(rsACADEMYget("buyname"))
		ELSE
			rsACADEMYget.close
			Exit function
		END IF

		rsACADEMYget.close



		'// 메일 발송
		dim oMail
		dim MailHTML



		set oMail = New MailCls

		oMail.MailType = mailType '메일 종류별 고정값 (mailLib2.asp 참고)
		oMail.MailTitles = mailTitle
		oMail.SenderMail = mailFrom
		oMail.SenderNm = nameFrom

		oMail.AddrType = "string"
		oMail.ReceiverNm = buyerName
		oMail.ReceiverMail = mailTo

		MailHTML= oMail.getMailTemplate()

		IF MailHTML="" Then
			response.write "메일발송 실패-템플릿 불러오기"
	    	'dbACADEMYget.close()	:	response.End
		End IF

		getInfo(vOrderSerial)

		'// 실제 메일에 정보 치환
		MailHTML = replace(MailHTML,"[$USER_NAME$]", AstarUserName(buyerName)) ' 주문자 이름
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' 주문번호
		MailHTML = replace(MailHTML,"[$$ORDERITEM_INFO_HTML$$]",getOrderItemInfo(vOrderSerial)) ' 주문상품 정보
		MailHTML = replace(MailHTML,"[$$PAY_REQ_INFO_HTML$$]", PayInfoHTML & ReqInfoHTML)	'결제 정보 / 배송지 정보

		oMail.MailConts = MailHTML

		'oMail.Send()
		oMail.Send_CDO()
		'oMail.Send_CDONT()

		SET oMail = nothing

		fcSendMail = MailHTML

End Function

Public Function fcSendMailItem(vOrderSerial, detailidx, mailFrom, nameFrom, mailTitle, mailType)

		'// 배송정보 & 메일정보

		dim strSQL

		dim buyerName, subTotalPrice, reqName , reqZipcode , reqAddress , reqPhone ,repHp , reqComment , regdate



		strSQL =" SELECT top 1 buyname,buyemail " &_
				" ,reqName,reqZipcode ,reqAddress ,reqPhone , reqhp , comment , regdate " &_
				" FROM [db_academy].[dbo].tbl_academy_order_master "  &_
				" WHERE orderserial = '" + vOrderSerial + "'"

		rsACADEMYget.Open strSQL,dbACADEMYget,1

		IF  not rsACADEMYget.EOF  THEN
			mailTo 		= rsACADEMYget("buyemail")
			buyerName  	= db2html(rsACADEMYget("buyname"))
			regdate  	= rsACADEMYget("regdate")
		ELSE
			rsACADEMYget.close
			Exit function
		END IF

		rsACADEMYget.close



		'// 메일 발송
		dim oMail
		dim MailHTML



		set oMail = New MailCls

		oMail.MailType = mailType '메일 종류별 고정값 (mailLib2.asp 참고)
		oMail.MailTitles = mailTitle
		oMail.SenderMail = mailFrom
		oMail.SenderNm = nameFrom

		oMail.AddrType = "string"
		oMail.ReceiverNm = buyerName
		oMail.ReceiverMail = mailTo

		MailHTML= oMail.getMailTemplate()

		IF MailHTML="" Then
			response.write "메일발송 실패-템플릿 불러오기"
	    	'dbACADEMYget.close()	:	response.End
		End IF

		getInfo(vOrderSerial)

		'// 실제 메일에 정보 치환
		MailHTML = replace(MailHTML,"[$USER_NAME$]", AstarUserName(buyerName)) ' 주문자 이름
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' 주문번호
		MailHTML = replace(MailHTML,"[$$ORDERITEM_INFO_HTML$$]",getOrderItemInfo(vOrderSerial)) ' 주문상품 정보
		MailHTML = replace(MailHTML,"[$$PAY_REQ_INFO_HTML$$]", PayInfoHTML & ReqInfoHTML)	'결제 정보 / 배송지 정보

		oMail.MailConts = MailHTML

		'oMail.Send()
		oMail.Send_CDO()
		'oMail.Send_CDONT()

		SET oMail = nothing

		fcSendMail = MailHTML

End Function

''// 주문 상품 정보
Function getOrderItemInfo(vOrderSerial)

	IF trim(vOrderSerial) ="" Then
		EXIT Function
	END IF

	dim Main_HTML,Sub_HTML
	Main_HTML =	"<table width=""550"" border=""0"" cellspacing=""0"" cellpadding=""0""> " &_
				"<tr> " &_
				"	<td style=""padding:0 0 7 0;""><img src=""http://fiximage.10x10.co.kr/web2008/mail/a01_text01.gif"" width=""346"" height=""18""></td> " &_
				"</tr> " &_
				"<tr> " &_
				"	<td> " &_
				"		<table width=""548"" border=""0"" cellspacing=""0"" cellpadding=""5""> " &_
				"		<tr> " &_
				"			<td> " &_
				"				<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"" style=""border-bottom:1px solid #dddddd""> " &_
				"				<tr> " &_
				"					<td valign=""bottom""> " &_
				"						<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""> " &_
				"						[$$ITEM_SUB$$] " &_
				"						</table> " &_
				"					</td> " &_
				"				</tr> " &_
				"				</table> " &_
				"			</td> " &_
				"		</tr> " &_
				"		<tr> " &_
				"			<td align=""right"" class=""eng11pxblack""> " &_
				"			상품 총 금액 : [$TOTAL_PRICE$] 원 + 총 배송비 : [$TOTAL_DELIVERY_PRICE$] 원 = 총 주문금액 : <span class=""red12pxb""> [$TOTAL_SUM_PRICE$] </span> 원 </span></td> " &_
				"			</td> " &_
				"		</tr> " &_
				"		</table> " &_
				"	</td> " &_
				"</tr> " &_
				"</table> "

	Sub_HTML="<!-- 반복 시작 --> " &_
			" <tr>" &_
			" <a href=""http://www.thefingers.co.kr/diyshop/shop_prd.asp?itemid=[$ITEM_ID$]"">" &_
			" 	<td bgcolor=""#FFFFFF"">" &_
			" 		<table width=""100%"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top: 1px solid #dddddd"">" &_
			" 		<tr>" &_
			" 			<td width=""260"" align=""right"" style=""border-right: 1px solid #dddddd"">" &_
			" 				<table width=""255"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"">" &_
			" 				<tr>" &_
			" 					<td width=""50"" valign=""bottom"">" &_
			" 						<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
			" 						<tr>" &_
			" 							<td><img src=""[$ITEM_IMAGE_URL$]"" width=""60"" ></td>" &_
			" 						</tr>" &_
			" 						</table>" &_
			" 					</td>" &_
			" 					<td style=""padding:5 "">([$ITEM_ID$]) [$ITEM_NAME$]</td>" &_
			" 				</tr>" &_
			" 				</table>" &_
			" 			</td>" &_
			" 			<td align=""center"">" &_
			" 				<table width=""100%"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"" >" &_
			" 				<tr>" &_
			" 					<td width=""60"" height=""35"" align=""center"" bgcolor=""#eeeeee""style=""border-bottom:1px solid #dddddd"" >수 량</td>" &_
			" 					<td width=""40"" bgcolor=""#FFFFFF"" style=""padding:0 5 0 5;border-bottom:1px solid #dddddd"">[$ITEM_QUANTITY$]<!-- 2이상 볼드처리--></td>" &_
			" 					<td width=""60"" align=""center"" bgcolor=""#eeeeee""  style=""padding:0 5 0 5;border-bottom:1px solid #dddddd"">판매가격</td>" &_
			" 					<td bgcolor=""#FFFFFF"" style=""padding:0 5 0 5;border-bottom:1px solid #dddddd"">" &_
			" 						<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
			" 						<tr><td>[$ITEM_PRICE$]</td></tr>" &_
			" 						</table>" &_
			" 					</td>" &_
			" 				</tr>" &_
			" 				<tr>" &_
			" 					<td align=""center"" bgcolor=""#eeeeee"">마일리지</td>" &_
			" 					<td style=""padding:5"">[$ITEM_MILEAGE$]</td>" &_
			" 					<td align=""center"" bgcolor=""#eeeeee"" class=""red"" style=""padding:5"">주문금액</td>" &_
			" 					<td style=""padding:5""><strong class=""black12px"">[$ITEM_SUM_PRICE$]</strong></td>" &_
			" 				</tr>" &_
			" 				</table>" &_
			" 			</td>" &_
			" 		</tr>" &_
			" 		</table>" &_
			" 	</td>" &_
			" </a>" &_
			" </tr>" &_
			" <!-- 반복 끝 -->"

		dim strSQL
		dim ItemID
		dim ItemName
		dim ItemOptionName
		dim ItemImage
		dim ItemNo
		dim itemCost
		dim ItemMileage

		dim BufItemName
		dim BufCost
		dim DeliveryCost
		dim TotalCost
		dim ItemHTML
		dim tmpItemHTML

        '// 주문상품 정보

		TotalCost = 0
		DeliveryCost = 0
		BufCost = 0

        strSQL =" SELECT a.itemid, c.itemname, a.itemoptionname, a.mileage,c.smallimage, a.itemno, a.isupchebeasong ,c.orgPrice, a.detailidx " &_
				" ,(case when a.itemid<>0 then a.reducedprice else 0 end) as itemcost " &_
				" ,(case when a.itemid=0 then a.itemcost else 0 end) as dlvcost " &_
				" FROM [db_academy].[dbo].tbl_academy_order_detail a " &_
				" LEFT JOIN [db_academy].[dbo].tbl_diy_item c " &_
				" 	on a.itemid = c.itemid " &_
				" WHERE a.orderserial = '" + vOrderSerial + "' " &_
				" and (a.cancelyn<>'Y') " &_
				" ORDER BY a.isupchebeasong asc "

        rsACADEMYget.Open strSQL,dbACADEMYget,2

		IF not rsACADEMYget.EOF Then
		rsACADEMYget.Movefirst
		Do until rsACADEMYget.eof

			'if (CLng(detailidx) = rsACADEMYget("detailidx")) then
				ItemID =CStr(rsACADEMYget("itemid")) '상품 코드
				ItemName = db2html(rsACADEMYget("itemname"))'상품명
				ItemOptionName =db2html(rsACADEMYget("itemoptionname")) '옵션명
				ItemImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(ItemID) & "/" & rsACADEMYget("smallimage")	'상품 이미지
				ItemNo = FormatNumber(rsACADEMYget("itemno"),0)	'수량

				ItemMileage	= FormatNumber(rsACADEMYget("mileage"),0)
				itemCost	= FormatNumber(rsACADEMYget("itemcost"),0)

				IF rsACADEMYget("dlvcost")<>0 and not isnull(rsACADEMYget("dlvcost")) THEN
					DeliveryCost = Cint(DeliveryCost) + CInt(rsACADEMYget("dlvcost"))
				End IF

				BufItemName = ItemName
				IF ItemOptionName<>"" Then
					BufItemName = BufItemName & "<br><font color=""blue"">["& ItemOptionName &"]</font>"
				End IF
				BufCost= FormatNumber(itemCost*ItemNo,0) '주문금액(수량x판매가)
				'saleprice	= '할인가
				'mileage	= '마일리지

				TotalCost = TotalCost + BufCost '총주문액
	        	IF ItemID<>0 Then
		        	tmpItemHTML = Sub_HTML
					tmpItemHTML = replace(tmpItemHTML,"[$ITEM_ID$]",ItemID)
					tmpItemHTML = replace(tmpItemHTML,"[$ITEM_NAME$]",BufItemName)
					tmpItemHTML = replace(tmpItemHTML,"[$ITEM_IMAGE_URL$]",ItemImage)
					tmpItemHTML = replace(tmpItemHTML,"[$ITEM_QUANTITY$]",ItemNo)
					tmpItemHTML = replace(tmpItemHTML,"[$ITEM_PRICE$]",itemCost)
					tmpItemHTML = replace(tmpItemHTML,"[$ITEM_MILEAGE$]",ItemMileage)
					tmpItemHTML = replace(tmpItemHTML,"[$ITEM_SUM_PRICE$]",BufCost)
	        	ItemHTML = ItemHTML & tmpItemHTML
	        	END IF
	        	ItemID="": ItemName="" 	: ItemOptionName="": ItemImage="": ItemNo="": itemCost="" : ItemMileage="" : BufItemName="": tmpItemHTML=""
			'end if

        	rsACADEMYget.movenext

        Loop
        ELSE
        	getOrderItemInfo=""
        	rsACADEMYget.close
        	Exit Function
        End if

        rsACADEMYget.close

        Main_HTML= replace(Main_HTML,"[$TOTAL_PRICE$]",FormatNumber(CStr(TotalCost),0))
        Main_HTML= replace(Main_HTML,"[$TOTAL_DELIVERY_PRICE$]",FormatNumber(CStr(DeliveryCost),0))
        Main_HTML= replace(Main_HTML,"[$TOTAL_SUM_PRICE$]",FormatNumber(CStr(TotalCost+DeliveryCost),0))
        getOrderItemInfo = replace(Main_HTML,"[$$ITEM_SUB$$]",ItemHTML)

End Function


'// 결제 정보

Dim PayInfoHTML

Dim ReqInfoHTML

Dim MailTo , MailTo_Nm , Uregdate

Function getInfo(vOrderSerial)

	dim strInfo_Html , strSQL

	dim PayMethod		'결제방법
	dim PayMethodName 	'결제방법명
	dim PayStatus		'결제상태
	dim SpendMileage 	'마일리지 사용액
	dim	TenCardSpend	'할인권 사용액
	dim AllAtDisPrice	'기타할인액(올엣)
	dim TotalPayPrice	'총 결제금액
	dim AccountNo		'입금계좌 정보

	dim ReqName		'받으시는분
	dim ReqPhone	'전화번호
	dim ReqHp		'핸드폰
	dim ReqZipCode	'우편번호
	dim ReqAddress	'배송주소
	dim ReqComment	'배송메모
    dim ReqZipAddr ''배송주소1 
    
	PayInfoHTML = ""
	ReqInfoHTML = ""
	MailTo = ""
	MailTo_Nm = ""
	Uregdate = ""

	strSQL =" SELECT Top 1 BuyName , BuyEmail , AccountDiv,AccountNo,SubTotalPrice , regdate" &_
			" , IsNULL(miletotalprice,0) as SpendMileage , IsNULL(tencardspend,0) as TenCardSpend , 0 as AllAtDiscountPrice " &_
			" , ReqName , ReqPhone , ReqHp , ReqZipCode , ReqZipAddr, (ReqZipAddr + ' ' + ReqAddress) as ReqAllAddress , Comment " &_
			" FROM [db_academy].[dbo].tbl_academy_order_master " &_
			" WHERE cancelyn='N' and orderserial = '"& vOrderSerial &"' "

	rsACADEMYget.open strSQL, dbACADEMYget,2

	IF not rsACADEMYget.eof THEN

		MailTo_Nm  	= db2html(rsACADEMYget("BuyName"))
		MailTo 		= db2html(rsACADEMYget("BuyEmail"))

		Uregdate	= rsACADEMYget("regdate")

		PayMethod 		= CStr(rsACADEMYget("AccountDiv"))
		AccountNo 		= rsACADEMYget("AccountNo")
		SpendMileage 	= FormatNumber(rsACADEMYget("SpendMileage"),0)
		TenCardSpend 	= FormatNumber(rsACADEMYget("TenCardSpend"),0)
		AllAtDisPrice 	= FormatNumber(rsACADEMYget("AllAtDiscountPrice"),0)
		TotalPayPrice 	= FormatNumber(rsACADEMYget("SubTotalPrice"),0)

		ReqName 	= rsACADEMYget("ReqName")
		ReqPhone 	= rsACADEMYget("ReqPhone")
		ReqHp 		= rsACADEMYget("ReqHp")
		ReqZipCode 	= rsACADEMYget("ReqZipCode")
		ReqAddress 	= rsACADEMYget("ReqAllAddress")
		ReqComment 	= rsACADEMYget("Comment")
		ReqZipAddr  = rsACADEMYget("ReqZipAddr")

		getInfo 	= 0 '정상

	ELSE
		getInfo 	= -1 '오류
		PayInfoHTML		=""
		ReqInfoHTML		=""

		rsACADEMYget.Close
		Exit Function

	End IF

	rsACADEMYget.Close

	'//=============  결제 정보 시작 ================//

	SELECT CASE PayMethod
		CASE "100" '신용카드
			PayMethodName="신용카드"
			PayStatus	="결제완료"
		CASE "80" ' 올엣카드
			PayMethodName="올엣카드"
			PayStatus	="결제완료"
		CASE "20" ' 실시간 계좌이체
			PayMethodName="실시간 계좌이체"
			PayStatus	="결제완료"
		CASE "7" ' 무통장 입금
			PayMethodName="무통장 입금"
			PayStatus	="입금전 상태"
		CASE ELSE
			PayMethodName=""
			PayStatus	="입금전 상태"
	END SELECT

	PayInfoHTML= ""&_
		" <table width=""550"" border=""0"" cellspacing=""0"" cellpadding=""0""> "&_
		" <tr> "&_
		" 	<td style=""padding:0 0 7 0;""><img src=""http://fiximage.10x10.co.kr/web2008/mail/a01_text02.gif"" width=""60"" height=""18""></td> "&_
		" </tr> "&_
		" <tr> "&_
		" 	<td> "&_
		" 		<table width=""548""  border=""0"" cellspacing=""0"" cellpadding=""0""> "&_
		" 		<tr> "&_
		" 			<td align=""center""> "&_
		" 				<table width=""548"" height=""92""  border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top:1px solid #dddddd""> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> 결제방법 </td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayMethodName &" </td> "&_
		"					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">결제상태</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayStatus &"</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">마일리지사용액</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& SpendMileage &" P </td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">할인권사용액</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& TenCardSpend &" 원</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">기타 할인액</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& AllAtDisPrice &" 원</td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">총 결제금액 </td> "&_
		" 					<td valign=""bottom"" class=""price"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong> "& TotalPayPrice &" 원</strong></td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "


	IF PayMethod = "7" THEN '무통장 입금
		PayInfoHTML= 	PayInfoHTML &_
		" 		 <!-- 무통장입금 --> "&_
		" 		<tr> "&_
		" 			<td align=""center""  style=""padding:5 0 0 0 ""> "&_
		" 				<table width=""548"" height=""31""  border=""0"" cellpadding=""0"" cellspacing=""0""style=""border-top:1px solid #dddddd""> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> 입금 은행 정보 </td> "&_
		" 					<td valign=""bottom"" class=""BIG_Black"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong>&nbsp;"& AccountNo &" </strong> (주)텐바이텐</td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "&_
		" 		<tr> "&_
		" 			<td align=""left"" class=""black11px"" style=""padding:10 15 0 15"">* 무통장입금 확인은 평일 매일 오전 10시, 오후 3시 두번 이루어지며 입금확인시 배송이 이루어집니다.<br> "&_
		" 			* 무통장주문 후 7일이 지날때까지 입금이 안되면 주문은 자동으로 취소됩니다. 일부 한정상품 주문시 유의하여 주시기 바랍니다.</td> "&_
		" 		</tr> "
	END IF

		PayInfoHTML= 	PayInfoHTML &_
		" 		</table> "&_
		" 	</td> "&_
		" </tr> "&_
		" </table>"

	PayInfoHTML = PayInfoHTML

	'//=============  결제 정보 끝 ================//

	'//=============  배송지 정보 시작 =================//

	ReqInfoHTML= ""&_
	"		<table align=""center"" width=""590"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:590px; font-family:'맑은고딕','Malgun Gothic','돋움', dotum, sans-serif; vertical-align:top;""> "&_
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">이름</th> "&_
	"				<td style=""padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">"& AstarUserName(ReqName) &"</td> "&_
	"			</tr> "&_ 
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">휴대폰</th> "&_
	"				<td style=""padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">"& AstarPhoneNumber(ReqHp) &"</td> "&_
	"			</tr> "&_
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">주소</th> "&_
	"				<td style=""padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">["& ReqZipCode &"] " & ReqZipAddr & " (이하 생략)" &"</td> "&_
	"			</tr> "&_
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 27px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">전화번호</th> "&_
	"				<td style=""padding:5px 10px 27px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">"& AstarPhoneNumber(ReqPhone) &"</td> "&_
	"			</tr> "&_
	"		</table> "
	ReqInfoHTML = ReqInfoHTML

''	"			<tr> "&_
''	"				<th style=""width:140px; padding:25px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:400; border-top:1px solid #ddd;"">요청사항</th> "&_
''	"				<td style=""padding:25px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:400; border-top:1px solid #ddd;"">"& ReqComment &"</td> "&_
''	"			</tr> "&_
	
	'//=============  배송지 정보 끝 =================//

End Function

Function fcSendMailFinish_Dlv_DIY(vOrderSerial,vMakerid)

		IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

		dim strHTML_MAIN,strHTML_Sub
		' 배송 주체별 HTML
		strHTML_MAIN ="" &_
			"[$ITEMHTMLTABLE$]"

		' 기본 상품 설명부분 HTML
		strHTML_Sub ="" &_
			"<tr>" &_
			"	<td style=""width:210px; padding:30px 30px 30px 20px; border-top:1px solid #ddd; vertical-align:top; text-align:left;"">" &_
			"		<img src=""[$ITEM_IMAGE_URL$]"" alt=""[$ITEM_NAME$]"" width=""210"" height=""140"" style=""width:210px; height:140px;"" /></td>" &_
			"	<td style=""padding:30px 20px 30px 0; vertical-align:top; border-top:1px solid #ddd;"">" &_
			"		<strong style=""display:block; width:370px; font-size:22px; color:#000; text-align:left; text-overflow:ellipsis; overflow:hidden; white-space:nowrap;"">[$ITEM_NAME$]</strong>[$ITEM_OPT_NAME$]" &_
			"		<span style=""font-size:20px; color:#666; text-align:left;"">[$ITEM_PRICE$]원<br>[$ITEM_QUANTITY$]</span>[$ITEM_DLV_STATUS$][$ITEM_DELIVERY_LINK$]" &_
			"	</td>" &_
			"</tr>"


        '주문 상품 정보
		dim strSQL
		dim ITIMG , ITNM , ITID , ITOPNM , ITNO , ITCOST
		dim DLVSTS, DLVLKTXT
		dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
		dim isNowDLV,isOtherDLV '지금 배송,같이주문한 상품

		tmpHTML="":NowHTML="":OtherHTML=""

		strSQL =" SELECT a.itemid, a.itemoptionname, c.smallimage, a.itemname,a.makerid ," &_
				" (c.cate_large + c.cate_mid + c.cate_small) as itemserial," &_
				" a.itemcost as sellcash, a.itemno, a.isupchebeasong, a.songjangdiv, replace(isnull(a.songjangno,''),'-','') as songjangno, a.currstate" &_
				" ,s.divname,s.findurl , c.basicimage" &_
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
				ITIMG = "http://image.thefingers.co.kr/diyitem/webimage/basic/" & GetImageSubFolderByItemid(rsACADEMYget("itemid")) & "/" & rsACADEMYget("basicimage")
				' 상품 코드
				ITID = rsACADEMYget("itemid")
				'--- 상품명
				ITNM = db2html(rsACADEMYget("itemname"))
				'--- 상품옵션명
				ITOPNM = db2html(rsACADEMYget("itemoptionname"))

				IF ITOPNM<>"" then
					ITOPNM = "<p style=""min-height:70px; margin:0; padding:7px 0; font-size:18px; color:#666; text-align:left; vertical-align:top;"">[" & ITOPNM & "]</p>"
				END IF
				'--- 상품수량 -- 수량별 style
				ITNO = Cstr(rsACADEMYget("itemno"))
				IF rsACADEMYget("itemno")>1 THEN
					ITNO = " X " & Cstr(rsACADEMYget("itemno")) & "개"
				Else
					ITNO = ""
				END If
				
				ITCOST	= FormatNumber(rsACADEMYget("sellcash"),0)

				'--- 배송상태 지정
					IF rsACADEMYget("currstate") = 7 THEN
						 DLVSTS = "<br>배송현황 : 출고완료"
					 ELSE
						 DLVSTS = "<br>배송현황 : 상품준비중"
					 END IF
				'--- 택배/송장 설정
				IF ((Not isnull(rsACADEMYget("songjangno"))) and  (rsACADEMYget("songjangno")<>"") ) THEN
					DLVLKTXT ="<br><a href=""" & db2html(rsACADEMYget("findurl")) & rsACADEMYget("songjangno") & """ target=""_blank""  class=""link_title"">" & db2html(rsACADEMYget("divname")) & " " & rsACADEMYget("songjangno") & "</a>"
				else
					DLVLKTXT ="-"
				end if
				tmpHTML = strHTML_Sub
				tmpHTML = replace(tmpHTML,"[$ITEM_IMAGE_URL$]",ITIMG)
				tmpHTML = replace(tmpHTML,"[$ITEM_ID$]",ITID)
				tmpHTML = replace(tmpHTML,"[$ITEM_NAME$]",ITNM)
				If ITOPNM <> "" Then '//옵션 있을때만
				    tmpHTML = replace(tmpHTML,"[$ITEM_OPT_NAME$]",ITOPNM)
			    else
			        tmpHTML = replace(tmpHTML,"[$ITEM_OPT_NAME$]","<br>") ''2016/12/01 추가 eastone
				End If 
				tmpHTML = replace(tmpHTML,"[$ITEM_PRICE$]",ITCOST)
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
			'ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text01.gif"" width=""79"" height=""18"" alt=""출고된 상품의 배송정보"">"
			NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
			'NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
		Else
			NowHTML= ""
		END IF

		IF OtherHTML<>"" and isOtherDLV THEN
			'ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text02.gif"" width=""193"" height=""18"" alt="" 같이 주문하신 상품 배송현황"">"
			OtherHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",OtherHTML)
			'OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
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
		oMail.MailTitles	 = "[더핑거스]주문하신 상품에 대한 핑거스 배송안내입니다!"
		oMail.SenderNm		 = "더핑거스"
		oMail.SenderMail	 = "customer@thefingers.co.kr"
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
		MailHTML = replace(MailHTML,"[$USER_NAME$]", AstarUserName(MailTo_Nm)) ' 주문자 이름
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' 주문번호
		MailHTML = replace(MailHTML,"[$REGDATE$]", formatdate(Uregdate,"0000.00.00-00:00")) ' 주문일자
		MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '출고된 상품 HTML
		MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'같이 주문한상품 HTML
		MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'배송지 정보 HTML

		oMail.MailConts = MailHTML

		'oMail.Send_Mailer()
		oMail.Send_CDO()
		'oMail.Send_CDONT()

		SET oMail = nothing

		'response.write "aaaaaaaaaaaaa" & MailHTML

End Function

function SendMiChulgoMail(idx)
    ''require /lectureadmin/lib/classes/jumun/misendcls.asp
    dim oneMisend
    dim strMailHTML,strMailTitle, contentsHtml
	strMailHTML = ""
	strMailTitle = "[더핑거스] 출고 지연 안내메일입니다."

    set oneMisend = new COldMiSend
    oneMisend.FRectDetailIDx = idx
    oneMisend.getOneOldMisendItem

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "더핑거스"
		oMail.SenderMail	= "customer@thefingers.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate
		''parsing
		strMailHTML = replace(strMailHTML,":ORDERSERIAL:",oneMisend.FOneItem.Forderserial)
		strMailHTML = replace(strMailHTML,":ITEMIMAGEURL:",oneMisend.FOneItem.Fsmallimage)
		strMailHTML = replace(strMailHTML,":ITEMID:",oneMisend.FOneItem.Fitemid)
		strMailHTML = replace(strMailHTML,":REGDATE:",formatdate(oneMisend.FOneItem.FRegdate,"0000.00.00-00:00"))
		strMailHTML = replace(strMailHTML,":ITEMNAME:",oneMisend.FOneItem.Fitemname)
		strMailHTML = replace(strMailHTML,":ITEMOPTIONNAME:",oneMisend.FOneItem.FItemoptionName)
		strMailHTML = replace(strMailHTML,":ITEMCNT:",oneMisend.FOneItem.Fitemcnt)
		strMailHTML = replace(strMailHTML,":COMPANYNAME:",oneMisend.FOneItem.FMakerId)
		'strMailHTML = replace(strMailHTML,":COMPANYNAME:",oneMisend.FOneItem.getDlvCompanyName)
		strMailHTML = replace(strMailHTML,":MAYSENDDATE:",oneMisend.FOneItem.FMisendipgodate)

		if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.<br>*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		else
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		end if

		if (oneMisend.FOneItem.FMisendipgodate<>"") then
    		if (oneMisend.FOneItem.FMisendReason="03") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "안녕하세요.   고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품이 발송이 지연될 예정입니다.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"
        			contentsHtml = contentsHtml & "쇼핑에 불편을 드린 점 진심으로 사과드리며, 기분좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.<br>"
    		    else
        		    contentsHtml = "안녕하세요.   고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내 메일입니다.<br>"
        			contentsHtml = contentsHtml & "아래 발송예정일에 발송될 예정이오며, 부득이한 사정으로 상품취소를 원하시는 경우,<br>"
        			contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"

    		    end if
    		elseif (oneMisend.FOneItem.FMisendReason="02") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "안녕하세요.  고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품은 주문 후 제작되는 상품으로<br>"
        			contentsHtml = contentsHtml & "일반상품과 달리 주문제작기간이 소요되는 상품입니다.<br>"
        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해드리오니,<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
    		    else
        		    contentsHtml = "안녕하세요.  고객님<br>"
        			contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내 메일입니다.<br>"
        			contentsHtml = contentsHtml & "아래와 같이 발송예정일을 안내해 드립니다.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "판매자가 상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			else
        			    contentsHtml = contentsHtml & "상품을 발송할 때까지 조금만 기다려 주시면 감사하겠습니다.<br>"
        			end if
    			end if
    	    elseif (oneMisend.FOneItem.FMisendReason="04") then
    	        oMail.MailTitles = "[더핑거스] 출고 예정 안내메일입니다."

    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"


    		    else
        		    contentsHtml = "안녕하세요.  고객님<br>"
                    contentsHtml = contentsHtml & "고객님께서 주문하신 상품의 출고안내메일입니다.<br>"
                    contentsHtml = contentsHtml & "주문하신 상품은 <strong>예약배송상품</strong>으로 아래 발송예정일에 발송될 예정이며,<br>"
                    contentsHtml = contentsHtml & "부득이한 사정으로 상품취소를 원하시는 경우,<br>"
                    contentsHtml = contentsHtml & "고객행복센터로 연락 부탁드립니다.<br>"


    			end if
    		end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML

		'oMail.Send_Mailer()
		oMail.Send_CDO
	End IF

    ''메모에 저장.
    'contentsHtml = replace(contentsHtml,"발송예정일","발송예정일("&oneMisend.FOneItem.FMisendipgodate&")")
	'call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing


end function



%>