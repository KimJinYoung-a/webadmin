<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/classes/order/bankacctcls.asp" -->
<!-- #include virtual="lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%
function SendMailPayDelay(orderserial,mailfrom)
        dim sql,discountrate,paymethod, i
        dim mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal, ttlsumHTML, ttSumsale

        mailtitle = "[�ٹ�����] �ֹ��� ���� �Ա�Ȯ��(���Ա�) �ȳ������Դϴ�"

        dim myorder
        set myorder = new COrderMaster
        myorder.FRectOrderserial = orderserial
        myorder.QuickSearchOrderMaster

        if (myorder.FOneItem.IsForeignDeliver) then
            myorder.getEmsOrderInfo
        end if

        dim myorderdetail
        set myorderdetail = new COrderMaster
        myorderdetail.FRectOrderserial = orderserial
		myorderdetail.FRectForMail = "Y"
        myorderdetail.QuickSearchOrderDetail

        if (myorder.FResultCount<1) then Exit function

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        'fileName = dirPath&"\\email_pay_delay.htm"
        fileName = dirPath&"\\email_new_paydelay.html"


        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)


		dim SpendMile, tencardspend
		dim IsForeighDeliver : IsForeighDeliver = false
        '�ֹ����� Ȯ��.---------------------------------------------------------------------------


        mailto = myorder.FOneItem.Fbuyemail
        paymethod = trim(myorder.FOneItem.Faccountdiv)


        if paymethod = "7" then    ' ������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�������Ա�")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�Ա��� ����")
        elseif paymethod = "100" then   ' �ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ſ�ī��")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "20" then   ' �ǽð���ü
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ǽð���ü")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "80" then   ' �þ�
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�þ�ī��")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "110" then   ' OKCashbag+�ſ�ī��
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "OKCashbag+�ſ�ī��")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        elseif paymethod = "400" then   ' �ڵ�������
            mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "�ڵ���")
            mailcontent = replace(mailcontent,":IPKUMSTATUS:", "�����Ϸ�")
        else
        	mailcontent = replace(mailcontent,":ACCOUNTDIVNAME:", "")
        end if

        if (paymethod<>"7") then
            mailcontent = ReplaceText(mailcontent,"(<!-----bankinfo------>)[\s\S]*(<!-----/bankinfo------>)","")
            mailcontent = ReplaceText(mailcontent,"(<!-----banknotiinfo------>)[\s\S]*(<!-----/banknotiinfo------>)","")
        end if

        IsForeighDeliver = myorder.FOneItem.IsForeignDeliver

        if (IsForeighDeliver) then
            mailcontent = replace(mailcontent,":REQHPORREQEMAIL:", "�̸���") ' ������ �̸���
            mailcontent = replace(mailcontent,":REQHP:", myorder.FOneItem.Freqemail) ' ������ ��ȭ��ȣ=>�̸��Ϸ�
            mailcontent = replace(mailcontent,":COUNTRYNAME:", myorder.FOneItem.FcountryNameEn) ' ����.
            mailcontent = replace(mailcontent,":REQZIPCODE:", myorder.FOneItem.FemsZipCode) ' ��ۿ����ȣ
        else
            mailcontent = replace(mailcontent,":REQHPORREQEMAIL:", "�޴�����ȣ") ' �޴�����ȣ
            mailcontent = replace(mailcontent,":REQHP:", myorder.FOneItem.Freqhp) ' ������ ��ȭ��ȣ
            mailcontent = replace(mailcontent,":REQZIPCODE:", myorder.FOneItem.Freqzipcode) ' ��ۿ����ȣ
            mailcontent = ReplaceText(mailcontent,"(<!-- foreigndelivery -->)[\s\S]*(<!--/foreigndelivery -->)","")
        end if

        mailcontent = replace(mailcontent,":BUYNAME:", myorder.FOneItem.Fbuyname) ' �ֹ��� �̸�
        mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' �ֹ���ȣ
        mailcontent = replace(mailcontent,":REQNAME:", myorder.FOneItem.Freqname) ' ������ �̸�
        mailcontent = replace(mailcontent,":REQALLADDRESS:", myorder.FOneItem.FreqZipaddr + " " + myorder.FOneItem.Freqaddress) ' ����ּ�
        mailcontent = replace(mailcontent,":REQPHONE:", myorder.FOneItem.Freqphone) ' ������ ��ȭ��ȣ

        mailcontent = replace(mailcontent,":BEASONGMEMO:", myorder.FOneItem.Fcomment) ' ��۸޸�


    	if (paymethod="110") then
    	    mailcontent = replace(mailcontent,":MAJORTOTALPRICE:", formatNumber(myorder.FOneItem.TotalMajorPaymentPrice,0) & " (�ſ�ī��:" &FormatNumber(myorder.FOneItem.TotalMajorPaymentPrice-myorder.FOneItem.FokcashbagSpend,0)& ",  OKCashbag:" &FormatNumber(myorder.FOneItem.FokcashbagSpend,0) &")") ' �����Ѿ�
    	else
    	    mailcontent = replace(mailcontent,":MAJORTOTALPRICE:", formatNumber(myorder.FOneItem.TotalMajorPaymentPrice,0)) ' �����Ѿ�
        end if

        mailcontent = replace(mailcontent,":ACCOUNTNO:", myorder.FOneItem.Faccountno) ' �Աݰ���

        if (myorder.FOneItem.FsumPaymentEtc<>0) then
            mailcontent = replace(mailcontent,":SPENDTENCASH:", FormatNumber(myorder.FOneItem.FsumPaymentEtc,0))
        else
            mailcontent = ReplaceText(mailcontent,"(<!-----spendtencash------>)[\s\S]*(<!-----/spendtencash------>)","")
        end if


		'�ֹ������� ���� Ȯ��.-----------------------------------------------------------------------------
itemHtml = itemHtml + "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; font-size:12px; font-family:dotum, '����', sans-serif; color:#707070;"">"&vbcrlf
itemHtml = itemHtml + "<tr>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; color:#707070; font-size:12px; line-height:12px; font-family:dotum, '����', sans-serif; text-align:center;"">��ǰ</th>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; color:#707070; font-size:12px; line-height:12px; text-align:center; font-family:dotum, '����', sans-serif;"">��ǰ�ڵ�</th>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:240px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; color:#707070; font-size:12px; line-height:12px; text-align:center; font-family:dotum, '����', sans-serif;"">��ǰ��[�ɼ�]</th>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; color:#707070; font-size:12px; line-height:12px; text-align:right; font-family:dotum, '����', sans-serif;"">�ǸŰ���</th>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:22px; height:44px; border-bottom:solid 1px #eaeaea; background:#f8f8f8; color:#707070; font-size:12px; font-family:dotum, '����', sans-serif;""></th>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:35px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; color:#707070; font-size:12px; line-height:12px; text-align:center; font-family:dotum, '����', sans-serif;"">����</th>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; color:#707070; font-size:12px; line-height:12px; text-align:right; font-family:dotum, '����', sans-serif;"">�ֹ��ݾ�</th>"&vbcrlf
itemHtml = itemHtml + "	<th style=""width:23px; border-bottom:solid 1px #eaeaea; background:#f8f8f8;""></th>"&vbcrlf
itemHtml = itemHtml + "</tr>"&vbcrlf

        for i=0 to myorderdetail.FResultCount-1
        	if myorderdetail.FItemList(i).FItemID <> 0 then
itemHtml = itemHtml + "<tr>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:50px; padding:6px 0; border-bottom:solid 1px #eaeaea;""><img src=""" &  myorderdetail.FItemList(i).FSmallImage & """ width=""50"" height=""50"" alt="""" /></td>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:100px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; color:#707070; font-size:11px; line-height:11px; font-family:dotum, '����', sans-serif;"">"& myorderdetail.FItemList(i).FItemID &"</td>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:240px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; color:#707070; font-size:11px; line-height:17px; font-family:dotum, '����', sans-serif;"">["&myorderdetail.FItemList(i).Fmakerid& "]<br /> " & myorderdetail.FItemList(i).FItemName
	if ( myorderdetail.FItemList(i).FItemOptionName <>"") then
itemHtml = itemHtml + "		["& myorderdetail.FItemList(i).FItemOptionName &"] "
	End if
itemHtml = itemHtml + "	</td>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:right; line-height:17px; font-family:dotum, '����', sans-serif; text-align:right;"">"&vbcrlf

if (myorderdetail.FItemList(i).Fissailitem = "Y") then
itemHtml = itemHtml + "		<span style=""margin:0; padding:6px 0; font-size:11px; line-height:16px; color:#707070; font-weight:bold; font-family:dotum, '����', sans-serif; text-decoration:line-through; text-align:right;"">"&FormatNumber(myorderdetail.FItemList(i).Forgitemcost,0)&"��</span>"&vbcrlf
itemHtml = itemHtml + "		<br /><span style=""margin:0; padding:6px 0; color:#dd5555; font-size:12px; line-height:16px; font-weight:bold; font-family:dotum, '����', sans-serif; text-align:right;"">"&FormatNumber(myorderdetail.FItemList(i).getItemcostCouponNotApplied,0)&"��</span>"&vbcrlf
else
    if (Not IsNull(myorderdetail.FItemList(i).Fitemcouponidx)) then
    itemHtml = itemHtml + "	<span style=""margin:0; padding:6px 0; font-size:11px; font-weight:bold; line-height:16px; color:#707070; font-family:dotum, '����', sans-serif; text-decoration:line-through; text-align:right;"">"&FormatNumber(myorderdetail.FItemList(i).FitemcostCouponNotApplied,0)&"��</span>"&vbcrlf
    else
    itemHtml = itemHtml + "	<span style=""margin:0; padding:0; font-weight:bold; color:#707070; font-size:12px; line-height:17px; font-family:dotum, '����', sans-serif; text-align:right;"">"&FormatNumber(myorderdetail.FItemList(i).FitemcostCouponNotApplied,0)&"��</span>"&vbcrlf
    end if
end if

if (Not IsNull(myorderdetail.FItemList(i).Fitemcouponidx)) then
    itemHtml = itemHtml + "	<br /><span style=""margin:0; padding:6px 0; color:#dd5555; font-size:11px; line-height:17px; text-align:right; font-family:dotum, '����', sans-serif;""><img src=""http://mailzine.10x10.co.kr/2017/ico_coupon.png"" alt=""��������"" style=""margin:0; vertical-align:-2px; padding-right:2px; font-size:11px; line-height:17px; text-align:right; font-family:dotum, '����', sans-serif;""/>" &FormatNumber(myorderdetail.FItemList(i).FItemCost,0)& "��</span>"&vbcrlf
end if
itemHtml = itemHtml + "	</td>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:22px; padding:6px 0; border-bottom:solid 1px #eaeaea;""></td>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:35px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:13px; line-height:13px; color:#707070; text-align:center; font-weight:bold; font-family:dotum, '����', sans-serif;"">" &myorderdetail.FItemList(i).FItemNo& "</td>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:85px; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, '����', sans-serif;"">"&vbcrlf
itemHtml = itemHtml + "		<span style=""margin:0; padding:0; font-weight:bold; color:#707070; font-size:12px; line-height:17px; font-family:dotum, '����', sans-serif; text-align:right;"">" &FormatNumber(myorderdetail.FItemList(i).FItemCost*myorderdetail.FItemList(i).FItemNo,0) & "��</span>"&vbcrlf
itemHtml = itemHtml + "	</td>"&vbcrlf
itemHtml = itemHtml + "	<td style=""width:23px; border-bottom:solid 1px #eaeaea;"">&nbsp;</td>"&vbcrlf
itemHtml = itemHtml + "</tr>"&vbcrlf
			end if
        next
itemHtml = itemHtml + "</table>"&vbcrlf

		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' �ֹ��������̺� �ֱ�
        mailcontent = itemHtmlTotal


		IF (myorder.FOneItem.Fmiletotalprice<>0) then
			ttSumsale = ttSumsale + myorder.FOneItem.Fmiletotalprice
		End If
		IF (myorder.FOneItem.Ftencardspend<>0) then
		    ttSumsale = ttSumsale + myorder.FOneItem.Ftencardspend
		end if
		if (myorder.FOneItem.Fallatdiscountprice + myorder.FOneItem.Fspendmembership<>0) then
			ttSumsale = ttSumsale + myorder.FOneItem.Fallatdiscountprice + myorder.FOneItem.Fspendmembership
		end if

		ttlsumHTML = ""
		ttlsumHTML = ttlsumHTML + "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">"&vbcrlf
		ttlsumHTML = ttlsumHTML + "<tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "	<td style=""border:solid 5px #eaeaea;"">"&vbcrlf
		ttlsumHTML = ttlsumHTML + "		<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">"&vbcrlf
		ttlsumHTML = ttlsumHTML + "		<tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<th style=""width:123px; height:45px; margin:0; padding:0; background:#f8f8f8; font-size:14px; line-height:14px; color:#707070; text-align:center; font-family:dotum, '����', sans-serif; font-weight:bold;"">���� �� �ݾ�</th>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<th style=""width:20px; height:45px; background:#f8f8f8;""></th>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<th style=""width:130px; height:45px; margin:0; padding:0; background:#f8f8f8; font-size:14px; line-height:14px; color:#707070; text-align:center; font-family:dotum, '����', sans-serif; font-weight:bold;"">��ۺ�</th>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<th style=""width:20px; height:45px; background:#f8f8f8;""></th>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<th style=""width:123px; height:45px; margin:0; padding:0; background:#f8f8f8; font-size:14px; line-height:14px; color:#707070; text-align:center; font-family:dotum, '����', sans-serif; font-weight:bold;"">���� �ݾ�</th>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<th style=""width:20px; height:45px; background:#f8f8f8;""></th>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<th style=""width:194px; height:45px; margin:0; padding:0; background:#f8f8f8; font-size:14px; line-height:14px; color:#707070; text-align:center; font-family:dotum, '����', sans-serif; font-weight:bold;"">�� �ֹ� �ݾ�</th>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "		</tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "		<tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<td style=""width:123px; height:68px; margin:0; padding:0; font-size:15px; line-height:15px; color:#000; text-align:center; font-family:verdana;""><span style=""margin:0; padding:0; font-size:15px; line-height:15px; color:#000; text-align:center; font-family:verdana; font-weight:bold;"">"& FormatNumber((myorder.FOneItem.FTotalSum-myorderdetail.BeasongPay),0) &"</span>��</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<td style=""width:20px; height:68px; margin:0; padding:0; font-size:15px; line-height:25px; font-weight:bold; vertical-align:middle; font-family:verdana;"">+</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<td style=""width:130px; height:68px; margin:0; padding:0; font-size:15px; line-height:15px; color:#000; text-align:center; font-family:verdana;""><span style=""margin:0; padding:0; font-size:15px; line-height:15px; color:#000; text-align:center; font-family:verdana; font-weight:bold;"">"& FormatNumber(myorderdetail.BeasongPay,0) &"</span>��</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<td style=""width:20px; height:68px; margin:0; padding:0; font-size:20px; line-height:20px; font-weight:bold; vertical-align:middle; font-family:verdana;"">-</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<td style=""width:123px; height:68px; margin:0; padding:0; font-size:15px; line-height:15px; color:#000; text-align:center; font-family:verdana;""><span style=""margin:0; padding:0; font-size:15px; line-height:15px; color:#000; text-align:center; font-family:verdana; font-weight:bold;"">"& FormatNumber(ttSumsale,0) &"</span>��</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<td style=""width:20px; height:68px; margin:0; padding:0; font-size:20px; line-height:20px; font-weight:bold; vertical-align:middle; font-family:verdana;"">=</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "			<td style=""width:194px; height:68px; margin:0; padding:0; font-size:24px; line-height:24px; color:#dd5555; text-align:center; font-family:verdana; font-weight:bold;""><span style=""margin:0; padding:0; font-size:24px; line-height:24px; color:#dd5555; text-align:center; font-family:verdana; font-weight:bold; font-family:verdana;"">"& FormatNumber(myorder.FOneItem.FsubtotalPrice,0) &"</span>��</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "		</tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "		</table>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "	</td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "</tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "<tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "	<td style=""padding-top:9px; text-align:right; font-size:11px; line-height:11px; color:#808080; font-family:dotum, '����', sans-serif;"">�������ϸ��� <span style=""color:#dd5555; font-weight:bold;"">"& FormatNumber(myorder.FOneItem.Ftotalmileage,0) &"P</span></td>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "</tr>"&vbcrlf
		ttlsumHTML = ttlsumHTML + "</table>"&vbcrlf
        mailcontent = replace(mailcontent,":ORDERPRICESUMMARY:", ttlsumHTML) ' �ֹ� �հ�ݾ�

        set myorder = Nothing
        set myorderDetail = Nothing

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end function


'call SendMailPayDelay("17080918279","�ٹ�����<customer@10x10.co.kr>")
call SendMailPayDelay("17081132083","�ٹ�����<customer@10x10.co.kr>")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->