
<%

'+----------------------------------------------------------------------------------------------------------------------+
'|                                        ��ü ��� ��ǰ ���� �߼�                                                      |
'+----------------------------------------------------+-----------------------------------------------------------------+
'|             �� �� ��                               |                          ��    ��                               |
'+----------------------------------------------------+-----------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer(orderserial,makerid) | ��� ���� �߼�(��ü��� ���)                                   |
'|                                                    | ��뿹 : fcSendMailFinish_DlvTEN('012012304','1293495006')      |
'+----------------------------------------------------+-----------------------------------------------------------------+



Function fcSendMailFinish_Dlv_Designer(vOrderSerial,vMakerid)

		IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

		dim strHTML_MAIN,strHTML_Sub
		' ��� ��ü�� HTML
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

		' �⺻ ��ǰ ����κ� HTML
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
				"					<td width=""60"" height=""35"" align=""center"">�� ��</td>" &_
				"					<td width=""60"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF"">[$ITEM_QUANTITY$]</td>" &_
				"					<td width=""60"" align=""center"" style=""padding:0 5 0 5;"">�����Ȳ</td>" &_
				"					<td class=""black12px"" style=""padding:0 5 0 5;"" bgcolor=""#FFFFFF""> [$ITEM_DLV_STATUS$]</td>" &_
				"				</tr>" &_
				"				<tr height=""1"">" &_
				"					<td colspan=""4"" align=""center"" bgcolor=""#dddddd""></td>" &_
				"				</tr>" &_
				"				<tr>" &_
				"					<td align=""center"">�����</td>" &_
				"					<td colspan=""3"" style=""padding:5"" bgcolor=""#FFFFFF""><strong class=""Information_font"">[$ITEM_DELIVERY_LINK$]</strong></td>" &_
				"				</tr>" &_
				"				</table>" &_
				"			</td>" &_
				"		</tr>" &_
				"		</table>" &_
				"	</td>" &_
				"</tr>"


        '�ֹ� ��ǰ ����
		dim strSQL
		dim ITIMG , ITNM , ITID , ITOPNM , ITNO
		dim DLVSTS, DLVLKTXT
		dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
		dim isNowDLV,isOtherDLV '���� ���,�����ֹ��� ��ǰ

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

				'--- ��ǰ�̹���
				ITIMG = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsACADEMYget("itemid")) & "/" & rsACADEMYget("smallimage")
				' ��ǰ �ڵ�
				ITID = rsACADEMYget("itemid")
				'--- ��ǰ��
				ITNM = db2html(rsACADEMYget("itemname"))
				'--- ��ǰ�ɼǸ�
				ITOPNM = db2html(rsACADEMYget("itemoptionname"))

				IF ITOPNM<>"" then
					ITNM = ITNM & "<br><font color=""blue"">[" & ITOPNM & "]</font>"
				END IF
				'--- ��ǰ���� -- ������ style
				ITNO = Cstr(rsACADEMYget("itemno"))
				IF rsACADEMYget("itemno")>1 THEN
					ITNO = "<strong>" & Cstr(rsACADEMYget("itemno")) & "</strong>"
				END IF

				'--- ��ۻ��� ����
					IF rsACADEMYget("currstate") = 7 THEN
						 DLVSTS = "<span class=""black12px"">���Ϸ�</span>"
					 ELSE
						 DLVSTS = "��ǰ�غ���"
					 END IF
				'--- �ù�/���� ����
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
			ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text01.gif"" width=""79"" height=""18"" alt=""���� ��ǰ�� �������"">"
			NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
			NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
		Else
			NowHTML= ""
		END IF

		IF OtherHTML<>"" and isOtherDLV THEN
			ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2008/mail/a03_text02.gif"" width=""193"" height=""18"" alt="" ���� �ֹ��Ͻ� ��ǰ �����Ȳ"">"
			OtherHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",OtherHTML)
			OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
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
		oMail.MailTitles	 = "[�ٹ�����]�ֹ��Ͻ� ��ǰ�� ���� �ٹ����� ��۾ȳ��Դϴ�!"
		'oMail.SenderNm		 = "�ٹ�����"
		'oMail.SenderMail	 = "customer@10x10.co.kr"
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
		MailHTML = replace(MailHTML,"[$USER_NAME$]", MailTo_Nm) ' �ֹ��� �̸�
		MailHTML = replace(MailHTML,"[$ORDER_SERIAL$]", vOrderSerial) ' �ֹ���ȣ
		MailHTML = replace(MailHTML,"[$$DELIVERY_ITEM_INFO$$]",NowHTML) '���� ��ǰ HTML
		MailHTML = replace(MailHTML,"[$$DELIVERY_OTHER_ITEM_INFO$$]",OtherHTML)	'���� �ֹ��ѻ�ǰ HTML
		MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",ReqInfoHTML)	'����� ���� HTML

		oMail.MailConts = MailHTML
		oMail.MailerMailGubun = 4		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
		'oMail.Send_CDO()
		'oMail.Send_CDONT()

		SET oMail = nothing

End Function


%>
