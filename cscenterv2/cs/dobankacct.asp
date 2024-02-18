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
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if orderserial <> "" then
	if checkNotValidHTML(orderserial) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
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

			'1. 사용마일리지 환원
			'2. 쿠폰 환원
			'3. 마일리지 재계산
			'4. 주문마일리지는 결재완료일때 계산되므로 여기서는 계산하지 않는다.
			sqlStr = "exec [db_user].[dbo].[usp_DIY_NoPayedOrder_MileageCouponCancelOne] '" & orderserial(i) & "' "
			rsget_CS.Open sqlStr,dbget_CS,1

			if Not rsget_CS.Eof then
				'response.write rsget_CS("result")
			end if
			rsget_CS.Close

			'5. 사용마일리지 환원이 이루어졌는지 검증
			'6. 쿠폰 환원이 이루어졌는지 검증
			'7. 마스터 취소
			'8. 한정수량 조정
			sqlStr = "exec [db_academy].[dbo].[usp_DIY_NoPayedOrder_CancelOne] '" & orderserial(i) & "' "
			rsget.Open sqlStr,dbget,1

			if Not rsget.Eof then
				if (rsget("result") <> 0) then
					response.write "<script>alert('취소가 실패하였습니다.');</script>"
					response.write "취소실패 : " & orderserial(i) & " - " & rsget("result")
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

				mailItems = Mid(mailcontent,InStr(mailcontent,"<!-- 상품목록 -->"),InstrRev(mailcontent,"<!-- 상품목록 -->") + 13 - InStr(mailcontent,"<!-- 상품목록 -->") )
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

				strPrice = "상품주문금액 " &  FormatNumber((oordermaster.FOneItem.Ftotalsum - oorderdetail.BeasongPay),0) & "원&nbsp;&nbsp; + &nbsp;&nbsp;배송비 " &  FormatNumber(oorderdetail.BeasongPay,0) & "원&nbsp;&nbsp; - &nbsp;&nbsp;마일리지 " &  FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) & "원&nbsp;&nbsp; - &nbsp;&nbsp;할인 " &  FormatNumber((oordermaster.FOneItem.Ftencardspend + oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership),0) & "원&nbsp;&nbsp; = &nbsp;&nbsp;<span style=""COLOR: #c3080a; TEXT-DECORATION: none; font-weight:bold;"">" &  FormatNumber(oordermaster.FOneItem.FsubtotalPrice,0) & "</span>원"

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
				mailtitle = "[텐바이텐] 주문에 대한 입금확인(미입금) 안내메일입니다"

				'rw oordermaster.FOneItem.Fbuyemail
				sendmail "텐바이텐<customer@10x10.co.kr>", oordermaster.FOneItem.Fbuyemail, mailtitle, mailcontent

				sqlStr = " insert into [db_temp].[dbo].tbl_bankmail_sendlist(orderserial)"
				sqlStr = sqlStr + " values('" & arrOrderSerial(k) & "')"
				'rw sqlStr
				dbget.execute(sqlStr)

				Set oordermaster = Nothing
				Set oorderdetail = Nothing

			End If
		Next
	End If

'	기존 메일발송 주석처리함
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
alert('발송 되었습니다.');
<% else %>
alert('저장 되었습니다.');
<% end if %>
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->