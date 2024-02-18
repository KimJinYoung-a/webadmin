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

'// ȫ�浿 => ȫ*��
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
			''3�̻�
			result = Left(userName,1) & "*" & Right(userName,1)
	End Select

	AstarUserName = result
end function

Public Function fcSendMail_OrderFinish(vOrderSerial)

		dim mailFrom, nameFrom, mailTitle, mailType

		mailFrom = "customer@thefingers.co.kr"
		nameFrom = "���ΰŽ�"
		mailTitle = "[���ΰŽ�] �ֹ��� �Ϸ�Ǿ����ϴ�."
		mailType = "6"							'���� mailLib2.asp

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

		'// ������� & ��������

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



		'// ���� �߼�
		dim oMail
		dim MailHTML



		set oMail = New MailCls

		oMail.MailType = mailType '���� ������ ������ (mailLib2.asp ����)
		oMail.MailTitles = mailTitle
		oMail.SenderMail = mailFrom
		oMail.SenderNm = nameFrom

		oMail.AddrType = "string"
		oMail.ReceiverNm = buyerName
		oMail.ReceiverMail = mailTo

		MailHTML= oMail.getMailTemplate()

		IF MailHTML="" Then
			response.write "���Ϲ߼� ����-���ø� �ҷ�����"
	    	'dbACADEMYget.close()	:	response.End
		End IF

		getInfo(vOrderSerial)

		'// ���� ���Ͽ� ���� ġȯ
		MailHTML = replace(MailHTML,"[$USER_NAME$]", AstarUserName(buyerName)) ' �ֹ��� �̸�
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' �ֹ���ȣ
		MailHTML = replace(MailHTML,"[$$ORDERITEM_INFO_HTML$$]",getOrderItemInfo(vOrderSerial)) ' �ֹ���ǰ ����
		MailHTML = replace(MailHTML,"[$$PAY_REQ_INFO_HTML$$]", PayInfoHTML & ReqInfoHTML)	'���� ���� / ����� ����

		oMail.MailConts = MailHTML

		'oMail.Send()
		oMail.Send_CDO()
		'oMail.Send_CDONT()

		SET oMail = nothing

		fcSendMail = MailHTML

End Function

Public Function fcSendMailItem(vOrderSerial, detailidx, mailFrom, nameFrom, mailTitle, mailType)

		'// ������� & ��������

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



		'// ���� �߼�
		dim oMail
		dim MailHTML



		set oMail = New MailCls

		oMail.MailType = mailType '���� ������ ������ (mailLib2.asp ����)
		oMail.MailTitles = mailTitle
		oMail.SenderMail = mailFrom
		oMail.SenderNm = nameFrom

		oMail.AddrType = "string"
		oMail.ReceiverNm = buyerName
		oMail.ReceiverMail = mailTo

		MailHTML= oMail.getMailTemplate()

		IF MailHTML="" Then
			response.write "���Ϲ߼� ����-���ø� �ҷ�����"
	    	'dbACADEMYget.close()	:	response.End
		End IF

		getInfo(vOrderSerial)

		'// ���� ���Ͽ� ���� ġȯ
		MailHTML = replace(MailHTML,"[$USER_NAME$]", AstarUserName(buyerName)) ' �ֹ��� �̸�
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' �ֹ���ȣ
		MailHTML = replace(MailHTML,"[$$ORDERITEM_INFO_HTML$$]",getOrderItemInfo(vOrderSerial)) ' �ֹ���ǰ ����
		MailHTML = replace(MailHTML,"[$$PAY_REQ_INFO_HTML$$]", PayInfoHTML & ReqInfoHTML)	'���� ���� / ����� ����

		oMail.MailConts = MailHTML

		'oMail.Send()
		oMail.Send_CDO()
		'oMail.Send_CDONT()

		SET oMail = nothing

		fcSendMail = MailHTML

End Function

''// �ֹ� ��ǰ ����
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
				"			��ǰ �� �ݾ� : [$TOTAL_PRICE$] �� + �� ��ۺ� : [$TOTAL_DELIVERY_PRICE$] �� = �� �ֹ��ݾ� : <span class=""red12pxb""> [$TOTAL_SUM_PRICE$] </span> �� </span></td> " &_
				"			</td> " &_
				"		</tr> " &_
				"		</table> " &_
				"	</td> " &_
				"</tr> " &_
				"</table> "

	Sub_HTML="<!-- �ݺ� ���� --> " &_
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
			" 					<td width=""60"" height=""35"" align=""center"" bgcolor=""#eeeeee""style=""border-bottom:1px solid #dddddd"" >�� ��</td>" &_
			" 					<td width=""40"" bgcolor=""#FFFFFF"" style=""padding:0 5 0 5;border-bottom:1px solid #dddddd"">[$ITEM_QUANTITY$]<!-- 2�̻� ����ó��--></td>" &_
			" 					<td width=""60"" align=""center"" bgcolor=""#eeeeee""  style=""padding:0 5 0 5;border-bottom:1px solid #dddddd"">�ǸŰ���</td>" &_
			" 					<td bgcolor=""#FFFFFF"" style=""padding:0 5 0 5;border-bottom:1px solid #dddddd"">" &_
			" 						<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
			" 						<tr><td>[$ITEM_PRICE$]</td></tr>" &_
			" 						</table>" &_
			" 					</td>" &_
			" 				</tr>" &_
			" 				<tr>" &_
			" 					<td align=""center"" bgcolor=""#eeeeee"">���ϸ���</td>" &_
			" 					<td style=""padding:5"">[$ITEM_MILEAGE$]</td>" &_
			" 					<td align=""center"" bgcolor=""#eeeeee"" class=""red"" style=""padding:5"">�ֹ��ݾ�</td>" &_
			" 					<td style=""padding:5""><strong class=""black12px"">[$ITEM_SUM_PRICE$]</strong></td>" &_
			" 				</tr>" &_
			" 				</table>" &_
			" 			</td>" &_
			" 		</tr>" &_
			" 		</table>" &_
			" 	</td>" &_
			" </a>" &_
			" </tr>" &_
			" <!-- �ݺ� �� -->"

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

        '// �ֹ���ǰ ����

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
				ItemID =CStr(rsACADEMYget("itemid")) '��ǰ �ڵ�
				ItemName = db2html(rsACADEMYget("itemname"))'��ǰ��
				ItemOptionName =db2html(rsACADEMYget("itemoptionname")) '�ɼǸ�
				ItemImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(ItemID) & "/" & rsACADEMYget("smallimage")	'��ǰ �̹���
				ItemNo = FormatNumber(rsACADEMYget("itemno"),0)	'����

				ItemMileage	= FormatNumber(rsACADEMYget("mileage"),0)
				itemCost	= FormatNumber(rsACADEMYget("itemcost"),0)

				IF rsACADEMYget("dlvcost")<>0 and not isnull(rsACADEMYget("dlvcost")) THEN
					DeliveryCost = Cint(DeliveryCost) + CInt(rsACADEMYget("dlvcost"))
				End IF

				BufItemName = ItemName
				IF ItemOptionName<>"" Then
					BufItemName = BufItemName & "<br><font color=""blue"">["& ItemOptionName &"]</font>"
				End IF
				BufCost= FormatNumber(itemCost*ItemNo,0) '�ֹ��ݾ�(����x�ǸŰ�)
				'saleprice	= '���ΰ�
				'mileage	= '���ϸ���

				TotalCost = TotalCost + BufCost '���ֹ���
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


'// ���� ����

Dim PayInfoHTML

Dim ReqInfoHTML

Dim MailTo , MailTo_Nm , Uregdate

Function getInfo(vOrderSerial)

	dim strInfo_Html , strSQL

	dim PayMethod		'�������
	dim PayMethodName 	'���������
	dim PayStatus		'��������
	dim SpendMileage 	'���ϸ��� ����
	dim	TenCardSpend	'���α� ����
	dim AllAtDisPrice	'��Ÿ���ξ�(�ÿ�)
	dim TotalPayPrice	'�� �����ݾ�
	dim AccountNo		'�Աݰ��� ����

	dim ReqName		'�����ôº�
	dim ReqPhone	'��ȭ��ȣ
	dim ReqHp		'�ڵ���
	dim ReqZipCode	'�����ȣ
	dim ReqAddress	'����ּ�
	dim ReqComment	'��۸޸�
    dim ReqZipAddr ''����ּ�1 
    
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

		getInfo 	= 0 '����

	ELSE
		getInfo 	= -1 '����
		PayInfoHTML		=""
		ReqInfoHTML		=""

		rsACADEMYget.Close
		Exit Function

	End IF

	rsACADEMYget.Close

	'//=============  ���� ���� ���� ================//

	SELECT CASE PayMethod
		CASE "100" '�ſ�ī��
			PayMethodName="�ſ�ī��"
			PayStatus	="�����Ϸ�"
		CASE "80" ' �ÿ�ī��
			PayMethodName="�ÿ�ī��"
			PayStatus	="�����Ϸ�"
		CASE "20" ' �ǽð� ������ü
			PayMethodName="�ǽð� ������ü"
			PayStatus	="�����Ϸ�"
		CASE "7" ' ������ �Ա�
			PayMethodName="������ �Ա�"
			PayStatus	="�Ա��� ����"
		CASE ELSE
			PayMethodName=""
			PayStatus	="�Ա��� ����"
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
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> ������� </td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayMethodName &" </td> "&_
		"					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">��������</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& PayStatus &"</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">���ϸ�������</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& SpendMileage &" P </td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">���αǻ���</td> "&_
		" 					<td valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& TenCardSpend &" ��</td> "&_
		" 				</tr> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">��Ÿ ���ξ�</td> "&_
		" 					<td width=""140"" valign=""bottom"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> "& AllAtDisPrice &" ��</td> "&_
		" 					<td width=""110"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd"">�� �����ݾ� </td> "&_
		" 					<td valign=""bottom"" class=""price"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong> "& TotalPayPrice &" ��</strong></td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "


	IF PayMethod = "7" THEN '������ �Ա�
		PayInfoHTML= 	PayInfoHTML &_
		" 		 <!-- �������Ա� --> "&_
		" 		<tr> "&_
		" 			<td align=""center""  style=""padding:5 0 0 0 ""> "&_
		" 				<table width=""548"" height=""31""  border=""0"" cellpadding=""0"" cellspacing=""0""style=""border-top:1px solid #dddddd""> "&_
		" 				<tr> "&_
		" 					<td width=""110"" height=""30"" align=""left"" valign=""bottom"" bgcolor=""#eeeeee""  style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""> �Ա� ���� ���� </td> "&_
		" 					<td valign=""bottom"" class=""BIG_Black"" style=""padding:0 0 6 10;border-bottom:1px solid #dddddd""><strong>&nbsp;"& AccountNo &" </strong> (��)�ٹ�����</td> "&_
		" 				</tr> "&_
		" 				</table> "&_
		" 			</td> "&_
		" 		</tr> "&_
		" 		<tr> "&_
		" 			<td align=""left"" class=""black11px"" style=""padding:10 15 0 15"">* �������Ա� Ȯ���� ���� ���� ���� 10��, ���� 3�� �ι� �̷������ �Ա�Ȯ�ν� ����� �̷�����ϴ�.<br> "&_
		" 			* �������ֹ� �� 7���� ���������� �Ա��� �ȵǸ� �ֹ��� �ڵ����� ��ҵ˴ϴ�. �Ϻ� ������ǰ �ֹ��� �����Ͽ� �ֽñ� �ٶ��ϴ�.</td> "&_
		" 		</tr> "
	END IF

		PayInfoHTML= 	PayInfoHTML &_
		" 		</table> "&_
		" 	</td> "&_
		" </tr> "&_
		" </table>"

	PayInfoHTML = PayInfoHTML

	'//=============  ���� ���� �� ================//

	'//=============  ����� ���� ���� =================//

	ReqInfoHTML= ""&_
	"		<table align=""center"" width=""590"" cellpadding=""0"" cellspacing=""0"" border=""0"" style=""width:590px; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; vertical-align:top;""> "&_
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">�̸�</th> "&_
	"				<td style=""padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">"& AstarUserName(ReqName) &"</td> "&_
	"			</tr> "&_ 
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">�޴���</th> "&_
	"				<td style=""padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">"& AstarPhoneNumber(ReqHp) &"</td> "&_
	"			</tr> "&_
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">�ּ�</th> "&_
	"				<td style=""padding:5px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">["& ReqZipCode &"] " & ReqZipAddr & " (���� ����)" &"</td> "&_
	"			</tr> "&_
	"			<tr> "&_
	"				<th style=""width:140px; padding:5px 10px 27px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">��ȭ��ȣ</th> "&_
	"				<td style=""padding:5px 10px 27px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:normal;"">"& AstarPhoneNumber(ReqPhone) &"</td> "&_
	"			</tr> "&_
	"		</table> "
	ReqInfoHTML = ReqInfoHTML

''	"			<tr> "&_
''	"				<th style=""width:140px; padding:25px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:400; border-top:1px solid #ddd;"">��û����</th> "&_
''	"				<td style=""padding:25px 10px 7px 10px; font-size:22px; color:#666; text-align:left; vertical-align:top; font-weight:400; border-top:1px solid #ddd;"">"& ReqComment &"</td> "&_
''	"			</tr> "&_
	
	'//=============  ����� ���� �� =================//

End Function

Function fcSendMailFinish_Dlv_DIY(vOrderSerial,vMakerid)

		IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

		dim strHTML_MAIN,strHTML_Sub
		' ��� ��ü�� HTML
		strHTML_MAIN ="" &_
			"[$ITEMHTMLTABLE$]"

		' �⺻ ��ǰ ����κ� HTML
		strHTML_Sub ="" &_
			"<tr>" &_
			"	<td style=""width:210px; padding:30px 30px 30px 20px; border-top:1px solid #ddd; vertical-align:top; text-align:left;"">" &_
			"		<img src=""[$ITEM_IMAGE_URL$]"" alt=""[$ITEM_NAME$]"" width=""210"" height=""140"" style=""width:210px; height:140px;"" /></td>" &_
			"	<td style=""padding:30px 20px 30px 0; vertical-align:top; border-top:1px solid #ddd;"">" &_
			"		<strong style=""display:block; width:370px; font-size:22px; color:#000; text-align:left; text-overflow:ellipsis; overflow:hidden; white-space:nowrap;"">[$ITEM_NAME$]</strong>[$ITEM_OPT_NAME$]" &_
			"		<span style=""font-size:20px; color:#666; text-align:left;"">[$ITEM_PRICE$]��<br>[$ITEM_QUANTITY$]</span>[$ITEM_DLV_STATUS$][$ITEM_DELIVERY_LINK$]" &_
			"	</td>" &_
			"</tr>"


        '�ֹ� ��ǰ ����
		dim strSQL
		dim ITIMG , ITNM , ITID , ITOPNM , ITNO , ITCOST
		dim DLVSTS, DLVLKTXT
		dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
		dim isNowDLV,isOtherDLV '���� ���,�����ֹ��� ��ǰ

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

				'--- ��ǰ�̹���
				ITIMG = "http://image.thefingers.co.kr/diyitem/webimage/basic/" & GetImageSubFolderByItemid(rsACADEMYget("itemid")) & "/" & rsACADEMYget("basicimage")
				' ��ǰ �ڵ�
				ITID = rsACADEMYget("itemid")
				'--- ��ǰ��
				ITNM = db2html(rsACADEMYget("itemname"))
				'--- ��ǰ�ɼǸ�
				ITOPNM = db2html(rsACADEMYget("itemoptionname"))

				IF ITOPNM<>"" then
					ITOPNM = "<p style=""min-height:70px; margin:0; padding:7px 0; font-size:18px; color:#666; text-align:left; vertical-align:top;"">[" & ITOPNM & "]</p>"
				END IF
				'--- ��ǰ���� -- ������ style
				ITNO = Cstr(rsACADEMYget("itemno"))
				IF rsACADEMYget("itemno")>1 THEN
					ITNO = " X " & Cstr(rsACADEMYget("itemno")) & "��"
				Else
					ITNO = ""
				END If
				
				ITCOST	= FormatNumber(rsACADEMYget("sellcash"),0)

				'--- ��ۻ��� ����
					IF rsACADEMYget("currstate") = 7 THEN
						 DLVSTS = "<br>�����Ȳ : ���Ϸ�"
					 ELSE
						 DLVSTS = "<br>�����Ȳ : ��ǰ�غ���"
					 END IF
				'--- �ù�/���� ����
				IF ((Not isnull(rsACADEMYget("songjangno"))) and  (rsACADEMYget("songjangno")<>"") ) THEN
					DLVLKTXT ="<br><a href=""" & db2html(rsACADEMYget("findurl")) & rsACADEMYget("songjangno") & """ target=""_blank""  class=""link_title"">" & db2html(rsACADEMYget("divname")) & " " & rsACADEMYget("songjangno") & "</a>"
				else
					DLVLKTXT ="-"
				end if
				tmpHTML = strHTML_Sub
				tmpHTML = replace(tmpHTML,"[$ITEM_IMAGE_URL$]",ITIMG)
				tmpHTML = replace(tmpHTML,"[$ITEM_ID$]",ITID)
				tmpHTML = replace(tmpHTML,"[$ITEM_NAME$]",ITNM)
				If ITOPNM <> "" Then '//�ɼ� ��������
				    tmpHTML = replace(tmpHTML,"[$ITEM_OPT_NAME$]",ITOPNM)
			    else
			        tmpHTML = replace(tmpHTML,"[$ITEM_OPT_NAME$]","<br>") ''2016/12/01 �߰� eastone
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
			'ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text01.gif"" width=""79"" height=""18"" alt=""���� ��ǰ�� �������"">"
			NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
			'NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
		Else
			NowHTML= ""
		END IF

		IF OtherHTML<>"" and isOtherDLV THEN
			'ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text02.gif"" width=""193"" height=""18"" alt="" ���� �ֹ��Ͻ� ��ǰ �����Ȳ"">"
			OtherHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",OtherHTML)
			'OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
		Else
			OtherHTML=""
		END IF


		'//=======  �������� & ������� , �������� �ҷ����� =========/
		call getInfo(vOrderSerial)

		IF MailTo ="" Then
			Exit Function
		End IF

		'//=======  ���� �߼� =========/
		dim oMail
		dim MailHTML

		set oMail = New MailCls

		oMail.MailType		 = 8 '���� ������ ������ (mailLib2.asp ����)
		oMail.MailTitles	 = "[���ΰŽ�]�ֹ��Ͻ� ��ǰ�� ���� �ΰŽ� ��۾ȳ��Դϴ�!"
		oMail.SenderNm		 = "���ΰŽ�"
		oMail.SenderMail	 = "customer@thefingers.co.kr"
		oMail.AddrType		 = "string"
		oMail.ReceiverNm	 = MailTo_Nm
		oMail.ReceiverMail	 = MailTo

		MailHTML= oMail.getMailTemplate()

		IF MailHTML="" Then
			SET oMail = nothing
			response.write "<script>alert('���Ϲ߼��� ���� �Ͽ����ϴ�.');</script>"
			Exit Function
	    End IF

		'// ���� ���Ͽ� ���� ġȯ
		MailHTML = replace(MailHTML,"[$USER_NAME$]", AstarUserName(MailTo_Nm)) ' �ֹ��� �̸�
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' �ֹ���ȣ
		MailHTML = replace(MailHTML,"[$REGDATE$]", formatdate(Uregdate,"0000.00.00-00:00")) ' �ֹ�����
		MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '���� ��ǰ HTML
		MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'���� �ֹ��ѻ�ǰ HTML
		MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'����� ���� HTML

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
	strMailTitle = "[���ΰŽ�] ��� ���� �ȳ������Դϴ�."

    set oneMisend = new COldMiSend
    oneMisend.FRectDetailIDx = idx
    oneMisend.getOneOldMisendItem

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "���ΰŽ�"
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
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.<br>*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		else
		    strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		end if

		if (oneMisend.FOneItem.FMisendipgodate<>"") then
    		if (oneMisend.FOneItem.FMisendReason="03") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.   ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� �߼��� ������ �����Դϴ�.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"
        			contentsHtml = contentsHtml & "���ο� ������ �帰 �� �������� ����帮��, ������� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.<br>"
    		    else
        		    contentsHtml = "�ȳ��ϼ���.   ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
        			contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"

    		    end if
    		elseif (oneMisend.FOneItem.FMisendReason="02") then
    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� �ֹ� �� ���۵Ǵ� ��ǰ����<br>"
        			contentsHtml = contentsHtml & "�Ϲݻ�ǰ�� �޸� �ֹ����۱Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��ص帮����,<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
    		    else
        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
        			contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>"
        			contentsHtml = contentsHtml & "�Ʒ��� ���� �߼ۿ������� �ȳ��� �帳�ϴ�.<br>"
        			if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
        			    contentsHtml = contentsHtml & "�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			else
        			    contentsHtml = contentsHtml & "��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>"
        			end if
    			end if
    	    elseif (oneMisend.FOneItem.FMisendReason="04") then
    	        oMail.MailTitles = "[���ΰŽ�] ��� ���� �ȳ������Դϴ�."

    		    if (oneMisend.FOneItem.getMisendDPlusDate>1) then
    		        contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"


    		    else
        		    contentsHtml = "�ȳ��ϼ���.  ����<br>"
                    contentsHtml = contentsHtml & "���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ������Դϴ�.<br>"
                    contentsHtml = contentsHtml & "�ֹ��Ͻ� ��ǰ�� <strong>�����ۻ�ǰ</strong>���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>"
                    contentsHtml = contentsHtml & "�ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>"
                    contentsHtml = contentsHtml & "���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>"


    			end if
    		end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML

		'oMail.Send_Mailer()
		oMail.Send_CDO
	End IF

    ''�޸� ����.
    'contentsHtml = replace(contentsHtml,"�߼ۿ�����","�߼ۿ�����("&oneMisend.FOneItem.FMisendipgodate&")")
	'call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing


end function



%>