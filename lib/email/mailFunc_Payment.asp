
<!-- #include virtual="/lib/email/mailFunction.asp" -->

<%

'+--------------------------------------------------------------------------------------------------------------------+
'|                                        �Ա� ���� ����                                                              |
'+--------------------------------------------------+-----------------------------------------------------------------+
'|             �� �� ��                             |                          ��    ��                               |
'+--------------------------------------------------+-----------------------------------------------------------------+
'| fcSendMail_PaymentInducement(orderserial)        | �Ա� ���� ���� �߼�()                                           |
'|                                                  | ��뿹 : fcSendMail_PaymentInducement('012012304')              |
'+--------------------------------------------------+-----------------------------------------------------------------+


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
			" 	<td bgcolor=""#FFFFFF"">" &_
			" 		<table width=""100%"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-top: 1px solid #dddddd"">" &_
			" 		<tr>" &_
			" 			<td width=""260"" align=""right"" style=""border-right: 1px solid #dddddd"">" &_
			" 				<table width=""255"" height=""70""  border=""0"" cellpadding=""0"" cellspacing=""0"">" &_
			" 				<tr>" &_
			" 					<td width=""50"" valign=""bottom"">" &_
			" 						<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">" &_
			" 						<tr>" &_
			" 							<td><img src=""[$ITEM_IMAGE_URL$]"" width=""50"" height=""50""></td>" &_
			" 						</tr>" &_
			" 						<tr>" &_
			" 							<td height=""17"" align=""center"" valign=""bottom"">[$ITEM_ID$]</td>" &_
			" 						</tr>" &_
			" 						</table>" &_
			" 					</td>" &_
			" 					<td style=""padding:5 "">[$ITEM_NAME$]</td>" &_
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
			" 						<!--<tr><td class=""price_line"">52,000</td></tr>" &_
			" 						<tr><td class=""black12px"">8,200��</td></tr>-->" &_
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
		
        strSQL =" SELECT a.itemid, c.itemname, a.itemoptionname, a.mileage,c.smallimage, a.itemno, a.isupchebeasong ,c.orgPrice " &_
				" ,(case when a.itemid<>0 then a.itemcost else 0 end) as itemcost " &_
				" ,(case when a.itemid=0 then a.itemcost else 0 end) as dlvcost " &_
				" FROM [db_order].[dbo].tbl_order_detail a " &_
				" JOIN [db_item].[dbo].tbl_item c " &_
				" 	on a.itemid = c.itemid " &_
				" WHERE a.orderserial = '" + vOrderSerial + "' " &_
				" and (a.cancelyn<>'Y') " &_
				" ORDER BY a.isupchebeasong asc "

        rsget.Open strSQL,dbget,2
		
		IF not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.eof
			
			ItemID =CStr(rsget("itemid")) '��ǰ �ڵ�
			ItemName = db2html(rsget("itemname"))'��ǰ��
			ItemOptionName =db2html(rsget("itemoptionname")) '�ɼǸ�
			ItemImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(ItemID) & "/" & rsget("smallimage")	'��ǰ �̹���
			ItemNo = FormatNumber(rsget("itemno"),0)	'����
			
			ItemMileage	= FormatNumber(rsget("mileage"),0)
			itemCost	= FormatNumber(rsget("itemcost"),0)
			
			IF rsget("dlvcost")<>0 and not isnull(rsget("dlvcost")) THEN
				DeliveryCost = Cint(DeliveryCost) + CInt(rsget("dlvcost"))
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
        	rsget.movenext
        Loop
        ELSE
        	getOrderItemInfo=""
        	rsget.close
        	Exit Function
        End if
        
        rsget.close
        
        Main_HTML= replace(Main_HTML,"[$TOTAL_PRICE$]",FormatNumber(CStr(TotalCost),0))
        Main_HTML= replace(Main_HTML,"[$TOTAL_DELIVERY_PRICE$]",FormatNumber(CStr(DeliveryCost),0))
        Main_HTML= replace(Main_HTML,"[$TOTAL_SUM_PRICE$]",FormatNumber(CStr(TotalCost+DeliveryCost),0))
        getOrderItemInfo = replace(Main_HTML,"[$$ITEM_SUB$$]",ItemHTML)
	
End Function


''// �Ա� ���� ����
Public Function fcSendMail_PaymentInducement(vOrderSerial)
		
		'// ������� & �������� 
			
		dim strSQL
	
		dim mailFrom , mailTo , mailTitle
		dim buyerName, subTotalPrice, reqName , reqZipcode , reqAddress , reqPhone ,repHp , reqComment
		
		
		
		strSQL =" SELECT top 1 buyname,buyemail " &_
				" ,reqName,reqZipcode ,reqAddress ,reqPhone , reqhp , comment " &_
				" FROM [db_order].[dbo].tbl_order_master "  &_
				" WHERE orderserial = '" + vOrderSerial + "'"
		
		rsget.Open strSQL,dbget,1
		
		IF  not rsget.EOF  THEN
			mailTo 		= rsget("buyemail")
			buyerName  	= db2html(rsget("buyname"))
		ELSE
			rsget.close
			Exit function
		END IF
		
		rsget.close
		
		IF mailTo ="" Then Exit Function 
		
		'// ���� �߼� 
		dim oMail
		dim MailHTML
		
		mailFrom = "customer@10x10.co.kr"
		mailTitle = "[�ٹ�����] �ֹ��� ���� �Ա�Ȯ��(���Ա�) �ȳ������Դϴ�"
		
		set oMail = New MailCls
		
		oMail.MailType = 9 '���� ������ ������ (mailLib2.asp ����)
		oMail.MailTitles = mailTitle
		oMail.SenderMail = mailFrom
		oMail.SenderNm = "�ٹ�����"
	
		oMail.AddrType = "string"
		oMail.ReceiverNm = buyerName
		oMail.ReceiverMail = mailTo
		
		MailHTML= oMail.getMailTemplate()
		
		IF MailHTML="" Then
			response.write "���Ϲ߼� ����-���ø� �ҷ�����"
	    	'dbget.close()	:	response.End
		End IF
		
		'// ���� ���Ͽ� ���� ġȯ
		MailHTML = replace(MailHTML,"[$USER_NAME$]", buyerName) ' �ֹ��� �̸�
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' �ֹ���ȣ
		MailHTML = replace(MailHTML,"[$$ORDERITEM_INFO_HTML$$]",getOrderItemInfo(vOrderSerial)) ' �ֹ���ǰ ���� 
		MailHTML = replace(MailHTML,"[$$PAY_INFO_HTML$$]",getPayInfo(vOrderSerial))	'���� ����
		MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",getReqInfo(vOrderSerial))	'����� ����
		
		oMail.MailConts = MailHTML
		
		'oMail.Send()
		oMail.Send_CDO()
		'oMail.Send_CDONT()
		
		SET oMail = nothing
		
		fcSendMail_PaymentInducement = MailHTML
	
End Function 


'fcSendMail_PaymentInducement("08092697465")

%>
