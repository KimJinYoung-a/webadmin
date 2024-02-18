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
dim mode,orderidx
mode = request("mode")
orderidx = Trim(request("orderidx"))

'response.write mode + "<br>"
'response.write orderidx + "<br>"

if Len(orderidx)>0 then
	if Right(orderidx,1)="," then orderidx=Left(orderidx,Len(orderidx)-1)
end if

'response.write orderidx
'dbget.close()	:	response.End

dim sqlStr,i
dim ibank
dim k, AssignedRow
set ibank = new CBankAcct

if mode="del" then
    ''취소일 추가
	sqlStr = "update [db_order].[dbo].tbl_order_master"
	sqlStr = sqlStr + " set cancelyn='Y'"
	sqlStr = sqlStr + " ,canceldate=getdate()"
	sqlStr = sqlStr + " where idx in (" + orderidx + ")"
	sqlStr = sqlStr + " and ipkumdiv='2'"
	sqlStr = sqlStr + " and accountdiv='7'"

	dbget.Execute sqlStr, AssignedRow
	k = AssignedRow

	'==========================================================================
    ''사용 마일리지 환급
	ibank.GetMileageSpendList orderidx
	for i=0 to ibank.FResultCount-1
'		response.write CStr(i) + "<br>"
		ibank.FItemList(i).DelMilegelog
		ibank.FItemList(i).RecalcuCurrentMileage
	next

	'==========================================================================
    ''사용 예치금 환급
	ibank.GetDepositSpendList orderidx
	for i=0 to ibank.FResultCount-1
		ibank.FItemList(i).DelDepositLog
		Call updateUserDeposit(ibank.FItemList(i).FUserid)
	next

	'==========================================================================
    ''사용 기프트카드 환급
	ibank.GetGiftCardSpendList orderidx
	for i=0 to ibank.FResultCount-1
		ibank.FItemList(i).DelGiftCardLog
		Call updateUserGiftCard(ibank.FItemList(i).FUserid)
	next

	'==========================================================================
    ''사용 쿠폰 환급
	ibank.GetCardSpendList orderidx
	for i=0 to ibank.FResultCount-1
'		response.write CStr(i) + "<br>"
		ibank.FItemList(i).DelCardSpend
	next


    ''재고 / 한정 업데이트
    sqlStr = "exec [db_summary].[dbo].sp_ten_RealtimeStock_cancelOrderAll_ByIdxARR '" & orderidx & "'"
    dbget.Execute(sqlStr)


	'// 삭제요청 목록에 전자보증서가 있으면 보증서 취소 요청 (2006.06.15; 운영관리팀 허진원)
	dim objUsafe, result, result_code, result_msg

	sqlStr =	"Select orderserial From db_order.[dbo].tbl_order_master " &_
				"Where idx in (" + orderidx + ") " &_
				"		and cancelyn='N' and InsureCd='0' "
	rsget.Open sqlStr,dbget,1

	Set objUsafe = CreateObject( "USafeCom.guarantee.1"  )

	if Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			' Real일 때
			objUsafe.Port = 80
			objUsafe.Url = "gateway.usafe.co.kr"
			objUsafe.CallForm = "/esafe/guartrn.asp"

			objUsafe.gubun	= "B0"				'// 전문구분 (A0:신규발급, B0:보증서취소, C0:입금확인)
			objUsafe.EncKey	= ""			'널값인 경우 암호화 안됨
			objUsafe.mallId	= "ZZcube1010"		'// 쇼핑몰ID
			objUsafe.oId	= CStr(rsget("orderserial"))	'// 주문번호

			'처리 실행!
			result = objUsafe.cancelInsurance

			rsget.MoveNext
		Loop
	end if

	Set objUsafe = Nothing

elseif mode="mail" Then

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

				call SendMailPayDelay(CStr(arrOrderSerial(k)), "텐바이텐<customer@10x10.co.kr>")

                sqlStr = " insert into [db_temp].[dbo].tbl_bankmail_sendlist(orderserial)"
				sqlStr = sqlStr + " values('" & arrOrderSerial(k) & "')"
				'rw sqlStr
				dbget.execute(sqlStr)

			End If
		Next
	End If

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
<!-- #include virtual="/lib/db/dbclose.asp" -->