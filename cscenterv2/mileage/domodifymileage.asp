<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<%

dim mode
dim userid, orderserial, mileage, jukyo
dim i, buf
'dim sqlStr


mode = requestCheckvar(request("mode"),32)
userid = requestCheckvar(request("userid"),32)
orderserial = requestCheckvar(request("orderserial"),32)
mileage = requestCheckvar(request("mileage"),32)
jukyo = requestCheckvar(request("jukyo"),32)



if (Not IsNumeric(mileage)) or (mileage="") then mileage = 0



if ((userid = "") or (orderserial = "")) then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if



'==============================================================================
''주문 마스타
dim oordermaster
set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster
end if

'' 과거 6개월 이전 내역 검색
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if



'==============================================================================
if ((orderserial <> "") and (oordermaster.FResultCount <> 1)) then
	response.write "<script>alert('잘못된 주문번호입니다.');</script>"
	response.write "<script>history.back();</script>"
	response.end

	orderserial = ""
end if



'==============================================================================
dim strSQL



dim divcd, reguserid, title, contents_jupsu, gubun01, gubun02
dim iAsID, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid

if (mode = "request") then
	'마일리지 적립요청

	if (IsExtSiteOrder(orderserial)) then
		divcd = "A005"		'외부몰 환불접수
	else
		divCd = "A003"		'마일리지 환불요청
	end if

	regUserID	= session("ssBctID")
	title = "마일리지 적립요청"
	contents_jupsu = jukyo
	gubun01		= "C004"	'공통
	gubun02		= "CD99"	'기타




	returnmethod = "R900"			'마일리지 환불
	refundrequire = CLng(mileage)	'환불 예정액
	orgsubtotalprice = 0
	orgitemcostsum = 0
	orgbeasongpay = 0
	orgmileagesum = 0
	orgcouponsum = 0
	orgallatdiscountsum = 0
	canceltotal = 0
	refunditemcostsum = 0
	refundmileagesum = 0
	refundcouponsum = 0
	allatsubtractsum = 0
	refundbeasongpay = 0
	refunddeliverypay = 0
	refundadjustpay = 0
	rebankname = ""
	rebankaccount = ""
	rebankownername = ""
	paygateTid = oordermaster.FOneItem.Fpaygatetid

	if IsNull(paygateTid) then
		paygateTid = ""
	end if

	iAsID = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	Call RegCSMasterRefundInfo(iAsID, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)


	response.write "<script>alert('요청 되었습니다.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>opener.focus(); window.close();</script>"

else
	'
end if

'response.write "aaaaaaaaaaa"

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->