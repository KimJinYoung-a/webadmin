<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/bankacctcls.asp"-->
<%
dim mode,orderidx, orderserial
mode = RequestCheckvar(request("mode"),16)
orderidx = Trim(request("orderidx"))
orderserial = Trim(request("orderserial"))
if orderidx <> "" then
	if checkNotValidHTML(orderidx) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if orderserial <> "" then
	if checkNotValidHTML(orderserial) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

'response.write mode + "<br>"
'response.write orderidx + "<br>"

orderidx = split(orderidx,"|")
orderserial = split(orderserial,"|")



dim sqlStr,i
dim ibank
dim k, AssignedRow
set ibank = new CBankAcct

if mode="del" then



	for i=0 to Ubound(orderserial)
		if orderserial(i)<>"" then

			'1. ��븶�ϸ��� ȯ��
			'2. ���� ȯ��
			'3. ���ϸ��� ����
			'4. �ֹ����ϸ����� ����Ϸ��϶� ���ǹǷ� ���⼭�� ������� �ʴ´�.
			sqlStr = "exec [db_user].[dbo].[usp_DIY_NoPayedOrder_MileageCouponCancelOne] '" & orderserial(i) & "' "
			rsget_CS.Open sqlStr,dbget_CS,1

			if Not rsget_CS.Eof then
				'response.write rsget_CS("result")
			end if
			rsget_CS.Close

			'5. ��븶�ϸ��� ȯ���� �̷�������� ����
			'6. ���� ȯ���� �̷�������� ����
			'7. ������ ���
			'8. �������� ����
			sqlStr = "exec [db_academy].[dbo].[usp_DIY_NoPayedOrder_CancelOne] '" & orderserial(i) & "' "
			rsget.Open sqlStr,dbget,1

			if Not rsget.Eof then
				if (rsget("result") <> 0) then
					response.write "<script>alert('��Ұ� �����Ͽ����ϴ�.');</script>"
					response.write "��ҽ��� : " & orderserial(i) & " - " & rsget("result")
					rsget.Close
					dbget.Close

					response.end
				end if
			end if
			rsget.Close

		end if
	next

elseif mode="mail" Then

	response.end

	Dim oordermaster, oorderdetail, strPrice
	Dim arrOrderSerial
	arrOrderSerial = Split(request("orderSerialArray"),",")
	'rw request("orderSerialArray")

	If IsArray(arrOrderSerial) Then

		dim fileContents, fs,dirPath,fileName,objFile, mailcontent, mailItems, tempItems, strItems
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.mappath("/lib/email")
		fileName = dirPath&"\\email_bank_resend.htm"
		Set objFile = fs.OpenTextFile(fileName,1)
		fileContents = objFile.readall

		For k = 0 To UBound(arrOrderSerial)
			If Len(arrOrderSerial(k)) = 11 Then

				mailcontent = fileContents

				mailItems = Mid(mailcontent,InStr(mailcontent,"<!-- ��ǰ��� -->"),InstrRev(mailcontent,"<!-- ��ǰ��� -->") + 13 - InStr(mailcontent,"<!-- ��ǰ��� -->") )
				mailcontent = Replace(mailcontent,mailItems,"|itemList|")

				Set objFile = Nothing
				Set fs = Nothing

				Set oordermaster = new COrderMaster
				oordermaster.FRectOrderSerial = arrOrderSerial(k)
				oordermaster.QuickSearchOrderMaster

				Set oorderdetail = new COrderMaster
				oorderdetail.FRectOrderSerial = arrOrderSerial(k)
				oorderdetail.QuickSearchOrderDetail

				strItems = ""
				For i = 0 To oorderdetail.FResultCount-1
					If oorderdetail.FItemList(i).Fitemid <> 0 Then
						tempItems = mailItems

						tempItems = Replace(tempItems, "|ItemID|", oorderdetail.FItemList(i).Fitemid)
						tempItems = Replace(tempItems, "|SmallImage|", oorderdetail.FItemList(i).FSmallImage)
						tempItems = Replace(tempItems, "|ItemName|", oorderdetail.FItemList(i).FItemName & "]" & oorderdetail.FItemList(i).FItemoptionName & "]")
						tempItems = Replace(tempItems, "|ItemNo|", oorderdetail.FItemList(i).FItemNo)

						tempItems = Replace(tempItems, "|ItemCost|", FormatNumber(oorderdetail.FItemList(i).FItemCost,0))
						tempItems = Replace(tempItems, "|ItemMileage|", FormatNumber(oorderdetail.FItemList(i).Fmileage * oorderdetail.FItemList(i).FItemNo,0))
						tempItems = Replace(tempItems, "|ItemPrice|", FormatNumber(oorderdetail.FItemList(i).FItemCost * oorderdetail.FItemList(i).FItemNo,0))

						strItems = strItems & tempItems


					End If
				Next

				strPrice = "��ǰ�ֹ��ݾ� " &  FormatNumber((oordermaster.FOneItem.Ftotalsum - oorderdetail.BeasongPay),0) & "��&nbsp;&nbsp; + &nbsp;&nbsp;��ۺ� " &  FormatNumber(oorderdetail.BeasongPay,0) & "��&nbsp;&nbsp; - &nbsp;&nbsp;���ϸ��� " &  FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) & "��&nbsp;&nbsp; - &nbsp;&nbsp;���� " &  FormatNumber((oordermaster.FOneItem.Ftencardspend + oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership),0) & "��&nbsp;&nbsp; = &nbsp;&nbsp;<span style=""COLOR: #c3080a; TEXT-DECORATION: none; font-weight:bold;"">" &  FormatNumber(oordermaster.FOneItem.FsubtotalPrice,0) & "</span>��"

				mailcontent = Replace(mailcontent, "|itemList|", strItems)
				mailcontent = Replace(mailcontent, "|strPrice|", strPrice)

			mailcontent = Replace(mailcontent, "|buyName|", oordermaster.FOneItem.FBuyName)
				mailcontent = Replace(mailcontent, "|orderSerial|", oordermaster.FOneItem.ForderSerial)

				mailcontent = Replace(mailcontent, "|JumunMethodName|", oordermaster.FOneItem.JumunMethodName)
				mailcontent = Replace(mailcontent, "|RegDate|", oordermaster.FOneItem.FRegDate)
				mailcontent = Replace(mailcontent, "|miletotalprice|", FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0))
				mailcontent = Replace(mailcontent, "|tencardspend|", FormatNumber(oordermaster.FOneItem.Ftencardspend,0))
				mailcontent = Replace(mailcontent, "|etcDiscount|", FormatNumber(oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership,0))
				mailcontent = Replace(mailcontent, "|subtotalPrice|", FormatNumber(oordermaster.FOneItem.FsubtotalPrice,0))
				mailcontent = Replace(mailcontent, "|accountno|", oordermaster.FOneItem.Faccountno)
				mailcontent = Replace(mailcontent, "|AccountName|", oordermaster.FOneItem.FAccountName)

				'rw mailcontent

				dim mailtitle
				mailtitle = "[�ٹ�����] �ֹ��� ���� �Ա�Ȯ��(���Ա�) �ȳ������Դϴ�"

				'rw oordermaster.FOneItem.Fbuyemail
				sendmail "�ٹ�����<customer@10x10.co.kr>", oordermaster.FOneItem.Fbuyemail, mailtitle, mailcontent

				sqlStr = " insert into [db_temp].[dbo].tbl_bankmail_sendlist(orderserial)"
				sqlStr = sqlStr + " values('" & arrOrderSerial(k) & "')"
				'rw sqlStr
				dbget.execute(sqlStr)

				Set oordermaster = Nothing
				Set oorderdetail = Nothing

			End If
		Next
	End If

'	���� ���Ϲ߼� �ּ�ó����
'	ibank.GetAcctRemailList orderidx
'	for i=0 to ibank.FResultCount-1
'		ibank.FItemList(i).AcctRemailSend
'	next

end if

set ibank = Nothing

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<% if (IsAutoScript) then %>
    <% response.write "S_OK|" & k %>
<% else %>
<script language="javascript">
<% if mode="mail" then %>
alert('�߼� �Ǿ����ϴ�.');
<% else %>
alert('���� �Ǿ����ϴ�.');
<% end if %>
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->