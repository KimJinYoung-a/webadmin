<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.08 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim i, masteridx, mode, divcd, orderno, ckAll ,IsOrderCanceled ,OrderMasterState ,ocsaslist
dim oordermaster , orgmasteridx , csmasteridx
	masteridx	= requestCheckVar(request("masteridx"),10)
	divcd		= requestCheckVar(request("divcd"),4)
	orderno	= requestCheckVar(request("orderno"),16)
	mode		= requestCheckVar(request("mode"),32)
	csmasteridx = requestCheckVar(request("csmasteridx"),10)
	
'CS접수마스터 가져오기
set ocsaslist = New COrder
	ocsaslist.FRectCsAsID = csmasteridx
	
	'/cs 마스터 테이블에 내역이 있는지 확인한다
	if (csmasteridx<>"") then
	    ocsaslist.fGetOneCSASMaster	   	    
	end if
	'response.write "<br>cs마스터테이블카운트:" & ocsaslist.ftotalcount & "!!<br>"
	
'CS접수마스터 정보가 없을경우 신규 접수
if (ocsaslist.FResultCount<1) then
	set ocsaslist.FOneItem = new COrderItem

	ocsaslist.FOneItem.fmasteridx = 0
	ocsaslist.FOneItem.Fdivcd = divcd

	mode = "regcsas"
else
    divcd       = ocsaslist.FOneItem.Fdivcd
    orderno = ocsaslist.FOneItem.Forderno
	masteridx = ocsaslist.FOneItem.forgmasteridx
	
    if (ocsaslist.FOneItem.FCurrState = "B007") then
    	mode = "finished"
    else
    	if (mode = "finishreginfo") then
    		'
    	else
    		mode = "editreginfo"
    	end if
    end if
end if

Call SetCSVariable_off(mode, divcd)

''주문 마스타
set oordermaster = new COrder
	oordermaster.FRectmasteridx = masteridx
	
	'/배송테이블 masteridx
	if (masteridx<>"") then
    	oordermaster.fQuickSearchOrderMaster
    end if

IsOrderCanceled = (oordermaster.FOneItem.Fcancelyn = "Y")
OrderMasterState = oordermaster.FOneItem.FIpkumDiv

'디테일 id(orderdetailidx)
dim distinctid

''접수 불가시 메세지
dim JupsuInValidMsg

if (Left(orderno,1)<>"A") and (oordermaster.ftotalcount<1) then
    response.write "<br><br>!!! 과거 주문내역이거나 주문 내역이 없습니다. - 관리자 문의 요망"
    dbget.close()	:	response.End
end if

''접수 가능 여부
dim IsJupsuProcessAvail
if (oordermaster.ftotalcount>0) then
    IsJupsuProcessAvail = ocsaslist.FOneItem.IsAsRegAvail_off(oordermaster.FOneItem.FIpkumdiv, oordermaster.FOneItem.FCancelyn , JupsuInValidMsg)
else
    IsJupsuProcessAvail = false
end if

'업체처리완료상태 여부
dim IsUpcheConfirmState
	IsUpcheConfirmState = (ocsaslist.FOneItem.FCurrState="B006")	
%>

<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript'>
	
var IsOrderMasterState			= <%=OrderMasterState %>;
var IsStatusRegister 			= <%= LCase(IsStatusRegister) %>;
var IsStatusEdit 				= <%= LCase(IsStatusEdit) %>;
var IsStatusFinishing 			= <%= LCase(IsStatusFinishing) %>;
var IsStatusFinished 			= <%= LCase(IsStatusFinished) %>;
var IsDisplayPreviousCSList 	= <%= LCase(IsDisplayPreviousCSList) %>;
var IsDisplayCSMaster 			= <%= LCase(IsDisplayCSMaster) %>;
var IsDisplayItemList 			= <%= LCase(IsDisplayItemList) %>;
var IsDisplayRefundInfo 		= <%= LCase(IsDisplayRefundInfo) %>;
var IsDisplayButton 			= <%= LCase(IsDisplayButton) %>;
var IsPossibleModifyCSMaster	= <%= LCase(IsPossibleModifyCSMaster) %>;
var IsPossibleModifyItemList	= <%= LCase(IsPossibleModifyItemList) %>;
var IsPossibleModifyRefundInfo	= <%= LCase(IsPossibleModifyRefundInfo) %>;
var IsDeletedCS 				= <%= LCase(ocsaslist.FOneITem.FDeleteyn = "Y") %>;

var CDEFAULTBEASONGPAY 		= "<%= getDefaultBeasongPayByDate(Left(Now, 10)) %>"; 	// 배송비

var divcd 					= "<%= divcd %>";
var mode 					= "<%= mode %>";
var orderno 			= "<%= orderno %>";
var IsAdminLogin 			= <%= LCase((session("ssBctId") = "icommang") or (session("ssBctId") = "iroo4") or (session("ssBctId") = "tozzinet")) %>;
var IsOrderFound 			= <%= LCase(oordermaster.ftotalcount > 0) %>;

<% if (oordermaster.ftotalcount > 0) then %>
	var IsThisMonthJumun 		= <%= LCase(datediff("m", oordermaster.FOneItem.FRegdate, now()) <= 0) %>;
<% else %>
	var IsThisMonthJumun 		= false;
<% end if %>

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<form name="popForm" action="/cscenter/ordermaster/popDeliveryTrace.asp" target="_blank">
	<input type="hidden" name="traceUrl">
	<input type="hidden" name="songjangNo">
</form>
<form name="frmaction" method="post" action="pop_cs_action_new_process.asp">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="modeflag2" value="<%= mode %>">
<input type="hidden" name="orderno" value="<%= oordermaster.FOneItem.forderno %>" >
<input type="hidden" name="csmasteridx" value="<%= ocsaslist.FOneItem.Fmasteridx %>">
<input type="hidden" name="masteridx" value="<%= oordermaster.FOneItem.fmasteridx %>">
<input type="hidden" name="divcd" value="<%= ocsaslist.FOneItem.FDivCd %>">
<input type="hidden" name="detailitemlist" value="">
<input type="hidden" name="ipkumdiv" value="<%= oordermaster.FOneItem.Fipkumdiv %>">
<input type="hidden" name="requireupche" value="">
<input type="hidden" name="requiremakerid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a">

<!-- 1. 이전 CS 내역                                                        -->
<!-- #include virtual="/admin/offshop/cscenter/action/inc_cs_action_prev_cslist.asp" -->

<!-- 2. CS 마스터 정보                                                      -->
<!-- #include virtual="/admin/offshop/cscenter/action/inc_cs_action_master_info.asp" -->

<!-- 3. 상품정보                                                            -->
<!-- #include virtual="/admin/offshop/cscenter/action/inc_cs_action_item_list.asp" -->
</table>

<!-- 5. 버튼                                                                -->
<!-- #include virtual="/admin/offshop/cscenter/action/inc_cs_action_button.asp" -->

</form>

<%
set oordermaster = Nothing
set ocsOrderDetail = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->