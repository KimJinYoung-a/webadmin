<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%

dim mode
dim userid, orderserial, mileage, jukyo, idx
dim gubun01, gubun02, gubun01name, gubun02name, contents_jupsu, requiremakerid
dim i, buf
'dim sqlStr


mode = requestCheckvar(request("mode"),32)
userid = requestCheckvar(request("userid"),32)
orderserial = requestCheckvar(request("orderserial"),32)
mileage = requestCheckvar(request("mileage"),32)
jukyo = requestCheckvar(request("jukyo"),32)
idx = requestCheckvar(request("idx"),32)

gubun01 = requestCheckvar(request("gubun01"),32)
gubun02 = requestCheckvar(request("gubun02"),32)
gubun01name = requestCheckvar(request("gubun01name"),32)
gubun02name = requestCheckvar(request("gubun02name"),32)
contents_jupsu = requestCheckvar(request("contents_jupsu"),2000)
requiremakerid = requestCheckvar(request("requiremakerid"),32)

if (Not IsNumeric(mileage)) or (mileage="") then mileage = 0



if ((userid = "") or ((orderserial = "") and (mode <> "requestForce") and (mode <> "delForce") and (mode <> "recalcmile"))) then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
''�ֹ� ����Ÿ
dim oordermaster
set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster
end if

'' ���� 6���� ���� ���� �˻�
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster

    'csAsFunction.asp
    GC_IsOldOrder = true
end if



'==============================================================================
if ((orderserial <> "") and (oordermaster.FResultCount <> 1)) then
	response.write "<script>alert('�߸��� �ֹ���ȣ�Դϴ�.');</script>"
	response.write "<script>history.back();</script>"
	response.end

	orderserial = ""
end if



'==============================================================================
dim strSQL



dim divcd, reguserid, title
dim iAsID, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid
dim isCSServiceRefund

if (mode = "request") then
	'���ϸ��� ������û

	isCSServiceRefund = False

	if (IsExtSiteOrder(orderserial)) then
		divcd = "A005"		'�ܺθ� ȯ������
	else
		divCd = "A003"		'���ϸ��� ȯ�ҿ�û
	end if

	regUserID	= session("ssBctID")

	if (gubun01 <> "") then
		title = "���ϸ��� ����(" & gubun02name & ")"
	else
		title = "���ϸ��� ����(" & jukyo & ")"
		contents_jupsu = jukyo
		gubun01		= "C004"	'����
		gubun02		= "CD99"	'��Ÿ
	end if

	returnmethod = "R900"			'���ϸ��� ȯ��
	refundrequire = CLng(mileage)	'ȯ�� ������
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

	Call SetCSServiceRefund(iAsID)
	if (requiremakerid <> "") and (requiremakerid <> "10x10logistics") then
		Call RegCSMasterAddUpche(iAsID, requiremakerid)
	end if

	response.write "<script>alert('��û �Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>opener.focus(); window.close();</script>"

elseif (mode = "requestForce") then

	regUserID	= session("ssBctID")

	if (gubun01 <> "") then
		title = "���ϸ��� ����(" & gubun02name & ")"
	else
		title = "���ϸ��� ����(" & jukyo & ")"
	end if

	strSQL = " insert into [db_user].[dbo].tbl_mileagelog(userid,mileage,jukyocd,jukyo,orderserial,regUserID)" & vbCrlf
	strSQL = strSQL + " values(" & vbCrlf
	strSQL = strSQL + " '" & userid & "'," & vbCrlf
	strSQL = strSQL + " " & mileage & "," & vbCrlf
	strSQL = strSQL + " '999'," & vbCrlf
	strSQL = strSQL + " '" & title & "'," & vbCrlf
	strSQL = strSQL + " ''," & vbCrlf
	strSQL = strSQL + " '" & regUserID & "'" & vbCrlf
	strSQL = strSQL + " )"
	dbget.Execute strSQL

	strSQL = "exec db_user.[dbo].[sp_Ten_ReCalcu_His_BonusMileage] '"& userid &"'"
	dbget.Execute strSQL

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>opener.focus(); window.close();</script>"

elseif (mode = "delForce") then

	strSQL = " update [db_user].[dbo].tbl_mileagelog " & vbCrlf
	strSQL = strSQL + " set deleteyn = 'Y' " & vbCrlf
	strSQL = strSQL + " 	where userid = '" & userid & "' and id = " & idx & " " & vbCrlf
	dbget.Execute strSQL

	strSQL = "exec db_user.[dbo].[sp_Ten_ReCalcu_His_BonusMileage] '"& userid &"'"
	dbget.Execute strSQL

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>history.back();</script>"

elseif (mode = "recalcmile") then

	strSQL = "exec db_user.[dbo].[sp_Ten_ReCalcu_His_BonusMileage] '"& userid &"'"
	dbget.Execute strSQL

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>history.back();</script>"

else
	'
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
