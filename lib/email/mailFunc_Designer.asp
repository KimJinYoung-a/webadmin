<!-- #include virtual="/lib/email/mailFunction.asp" -->
<%
'+--------------------------------------------------------------------------------------------------------------------------------+
'|                                        ��ü ��� ��ǰ ���� �߼�                                                                |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'|             �� �� ��                               |                          ��    ��                                         |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer(orderserial,makerid) | ��� ���� �߼�(��ü��� ���)                                             |
'|                                                    | ��뿹 : fcSendMailFinish_Dlv_Designer('012012304','1293495006')          |
'+----------------------------------------------------+---------------------------------------------------------------------------+
'| fcSendMailFinish_Dlv_Designer_off(detailidx,makerid)  �������� ��� ���� �߼�(��ü��� ���)                                   |
'|                                                    | ��뿹 : fcSendMailFinish_Dlv_Designer_off('012012304','1293495006')      |
'+----------------------------------------------------+---------------------------------------------------------------------------+

'' �ش�귣�� ��ü ���� �߼�(�ߺ� �߼� ����) ''//2014/03/31 �߰� , 2019/06/27 �ֱ� 1�ð��� �߼۰��� �ִ°�츸.
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

Function fcSendMailFinish_Dlv_Designer(vOrderSerial,vMakerid)	'/2011.04.21 �ѿ�� ����

	IF trim(vOrderSerial) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub ,strHTML_MAINother
	' ��� ��ü�� HTML
	strHTML_MAIN = ""
	strHTML_MAIN = strHTML_MAIN &"<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"	<td style=""padding:0 29px 45px; margin:0;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"			<th style=""margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, '����', sans-serif; text-align:left; color:#000;"">�߼۵� ��ǰ <span style=""margin-left:15px; padding:0; font-size:11px; line-height:11px; font-weight:normal; font-family:dotum, '����', sans-serif; vertical-align:2px; color:#808080; text-align:left;"">����� ��ȣ�� Ŭ���Ͻø� �����Ȳ�� Ȯ���Ͻ� �� �ֽ��ϴ�.</span><th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		</tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"		<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"			<td style=""border-top:solid 2px #000;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"				<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; font-size:12px; font-family:dotum, '����', sans-serif; color:#707070;"">"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"					<tr>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, '����', sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;"">��ǰ</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">��ǰ�ڵ�</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:250px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">��ǰ��[�ɼ�]</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:37px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">����</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:95px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">�ֹ�����</th>"&vbcrlf
	strHTML_MAIN = strHTML_MAIN &"						<th style=""width:108px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">�ù�����</th>"&vbcrlf
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
	strHTML_MAINother = strHTML_MAINother & "			<th style=""margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, '����', sans-serif; text-align:left; color:#000;"">�Բ� �ֹ��Ͻ� ��ǰ �����Ȳ<th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		</tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		<tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "			<td style=""border-top:solid 2px #000;"">"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "				<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; font-size:12px; font-family:dotum, '����', sans-serif; color:#707070;"">"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "					<tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, '����', sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;"">��ǰ</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">��ǰ�ڵ�</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:250px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">��ǰ��[�ɼ�]</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:37px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">����</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:95px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">�ֹ�����</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "						<th style=""width:108px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px; line-height:12px;"">�ù�����</th>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "					</tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "				[$ITEMHTMLTABLE$]"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "				</table>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "			</td>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "		</tr>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother &"		</table>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "	</td>"&vbcrlf
	strHTML_MAINother = strHTML_MAINother & "</tr>"

	' �⺻ ��ǰ ����κ� HTML
	strHTML_Sub =""
	strHTML_Sub = strHTML_Sub & "<tr>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:50px; padding:6px 0;border-bottom:solid 1px #eaeaea;"">"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		<img src=""[$ITEM_IMAGE_URL$]"" width=50 height=50 alt="""" />"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:100px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; color:#707070; font-size:11px; line-height:11px; font-family:dotum, '����', sans-serif;"">"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		[$ITEM_ID$]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:250px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; color:#707070; font-size:11px; line-height:17px; font-family:dotum, '����', sans-serif;"">"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		[[$ITEM_brandName$]]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "		<br /> [$ITEM_NAME$]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:37px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '����', sans-serif; color:#707070; font-size:13px; line-height:13px;"">[$ITEM_QUANTITY$]</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	[$ITEM_DLV_STATUS$]"&vbcrlf
	strHTML_Sub = strHTML_Sub & "	<td style=""width:108px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:12px; text-align:center; font-family:dotum, '����', sans-serif; text-align:center;"">[$ITEM_DELIVERY_LINK$]</td>"&vbcrlf
	strHTML_Sub = strHTML_Sub & "</tr>"

    '�ֹ� ��ǰ ����
	dim strSQL
	dim ITIMG , ITNM , ITID , ITOPNM , ITNO , ITbrandName ,ITmakerid
	dim DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '���� ���,�����ֹ��� ��ǰ

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

			'-- �귣��
			ITmakerid = db2html(rsget("makerid"))

			'-- �귣���
			ITbrandName = db2html(rsget("brandName"))

			'--- ��ǰ�̹���
			ITIMG = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("smallimage")
			' ��ǰ �ڵ�
			ITID = rsget("itemid")
			'--- ��ǰ��
			ITNM = db2html(rsget("itemname"))
			'--- ��ǰ�ɼǸ�
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & " [" & ITOPNM & "]"
			END IF
			'--- ��ǰ���� -- ������ style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = Cstr(rsget("itemno"))
			END IF

			'--- ��ۻ��� ����
				IF rsget("currstate") = 7 THEN
					DLVSTS = "<td style=""width:95px; height:44px; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '����', sans-serif; color:#dd5555; font-size:12px;"">���Ϸ�</td>"
				 ELSE
					DLVSTS = "<td style=""width:95px; height:44px; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, '����', sans-serif; color:#707070; font-size:12px;"">��ǰ�غ���</td>"
				 END IF
			'--- �ù�/���� ����
			IF ((Not isnull(rsget("songjangno"))) and  (rsget("songjangno")<>"") ) THEN
				DLVLKTXT = ""
				DLVLKTXT = DLVLKTXT & "<span style=""margin:0; padding:0; color:#707070; font-size:12px; font-weight:bold; line-height:18px; font-family:dotum, '����', sans-serif; text-align:center;"">" & db2html(rsget("divname")) & "</span><br />"
				DLVLKTXT = DLVLKTXT & "<a href=""" & db2html(rsget("findurl")) & rsget("songjangno") & """ style=""margin:0; padding:0; font-size:12px; color:#dd5555; font-size:11px; line-height:18px; font-family:dotum, '����', sans-serif; color:#0066cc; text-align:center;"">" & rsget("songjangno") & "</a>"
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
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_shiped.gif"" alt=""���� ��ǰ�� �������"">"
		NowHTML = replace(strHTML_MAIN,"[$ITEMHTMLTABLE$]",NowHTML)
		NowHTML = replace(NowHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		NowHTML= ""
	END IF

	IF OtherHTML<>"" and isOtherDLV THEN
		ITTITLEIMG ="<img src=""http://fiximage.10x10.co.kr/web2011/mail/tit_otherpd.gif"" alt="" ���� �ֹ��Ͻ� ��ǰ �����Ȳ"">"
		OtherHTML = replace(strHTML_MAINother,"[$ITEMHTMLTABLE$]",OtherHTML)
		OtherHTML = replace(OtherHTML,"[$DELIVERY_HOST_IMG$]",ITTITLEIMG)
	Else
		OtherHTML=""
	END IF


	'//=======  �������� & ������� , �������� �ҷ����� =========/
	'// ( !!!!! /lib/email/mailFunction.asp ���� !!!!! )
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
	MailHTML = replace(MailHTML,"[$$REQ_INFO_HTML$$]",newReqInfoHTML)	'����� ���� HTML
	MailHTML = replace(MailHTML,"http://mailzine.10x10.co.kr/2017/txt_noti_send_prd.png", "http://mailzine.10x10.co.kr/2017/txt_noti_send_prd2.png")	' �����ϰ�� ��� �̹��� ����. ��ü�� ����� ������� �ʰ� ������ ���̽��� ����.

	oMail.MailConts = MailHTML

	'response.write MailHTML
	'response.end
	oMail.MailerMailGubun = 4		' ���Ϸ� �ڵ����� ��ȣ
	oMail.Send_TMSMailer()		'TMS���Ϸ�
	'oMail.Send_Mailer()
	''oMail.Send_CDO()
	'oMail.Send_CDONT()

	SET oMail = nothing

End Function

Function fcSendMailFinish_Dlv_Designer_off(vmasteridx,vMakerid)
	IF trim(vmasteridx) ="" or vMakerid="" then EXIT Function

	dim strHTML_MAIN,strHTML_Sub

	dim vOrderSerial

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

	' �⺻ ��ǰ ����κ� HTML '�̹���"<td><img src=""[$ITEM_IMAGE_URL$]"" width=""50"" height=""50""></td>" &_
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
	dim strSQL, ITIMG , ITNM , ITID , ITOPNM , ITNO ,DLVSTS, DLVLKTXT
	dim tmpHTML,NowHTML,OtherHTML,ITTITLEIMG
	dim isNowDLV,isOtherDLV '���� ���,�����ֹ��� ��ǰ

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

			'--- ��ǰ�̹���
			ITIMG = ""
			' ��ǰ �ڵ�
			ITID = rsget("itemgubun")&Format00(6,rsget("itemid"))&rsget("itemoption")
			'--- ��ǰ��
			ITNM = db2html(rsget("itemname"))
			'--- ��ǰ�ɼǸ�
			ITOPNM = db2html(rsget("itemoptionname"))

			IF ITOPNM<>"" then
				ITNM = ITNM & "<br><font color=""blue"">[" & ITOPNM & "]</font>"
			END IF
			'--- ��ǰ���� -- ������ style
			ITNO = Cstr(rsget("itemno"))
			IF rsget("itemno")>1 THEN
				ITNO = "<strong>" & Cstr(rsget("itemno")) & "</strong>"
			END IF

			'--- ��ۻ��� ����
				IF rsget("currstate") = 7 THEN
					 DLVSTS = "<span class=""black12px"">���Ϸ�</span>"
				 ELSE
					 DLVSTS = "��ǰ�غ���"
				 END IF
			'--- �ù�/���� ����
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
	call getInfo_off(vmasteridx)

	IF MailTo ="" Then
		Exit Function
	End IF

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	oMail.MailType		 = 8 '���� ������ ������ (mailLib2.asp ����)
	oMail.MailTitles	 = "[�ٹ����ټ�]�ֹ��Ͻ� ��ǰ�� ���� �ٹ����� ��۾ȳ��Դϴ�!"
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
