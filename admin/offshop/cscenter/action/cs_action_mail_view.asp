<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.09 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_mail_Function_off.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim masteridx, ForceCurrState, ForceBuyEmail, resend ,oCsAction,strHTML, orgBuyEmail, orgCurrState
	masteridx = RequestCheckVar(request("masteridx"),10)
	ForceCurrState  = RequestCheckVar(request("ForceCurrState"),10)
	resend          = RequestCheckVar(request("resend"),10)

Set oCsAction = New CsActionMailCls
	if (ForceCurrState<>"") then oCsAction.FRectForceCurrState = ForceCurrState
	if (ForceBuyEmail<>"") then oCsAction.FRectForceBuyEmail = ForceBuyEmail
	
	strHTML = oCsAction.makeMailTemplate_off(masteridx)
	
	orgBuyEmail     = oCsAction.FBuyEmail
	orgCurrState    = oCsAction.FCurrState
	
	if (ForceBuyEmail<>"") then orgBuyEmail = ForceBuyEmail
	if (ForceCurrState<>"") then orgCurrState = ForceCurrState
	
	
	if (resend="on") then
	    call ReSendCsActionMail_off(masteridx, ForceCurrState, ForceBuyEmail)
	end if

Set oCsAction = Nothing
%>

<script language='javascript'>

function jsReloadMe(){
    var frm = document.submitFrm;
    frm.resend.value = "";
    frm.submit();
}

function ReSendMail(){
    var frm = document.submitFrm;
    if (confirm('메일을 재발송 하시겠습니까?')){
        frm.resend.value = "on"
        frm.submit();
    }
}

</script>

<table width="600" border=0 cellspacing=0 cellpadding=5 align="center">
<form name=submitFrm method=get action="">
<input type="hidden" name="resend" value="">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<tr>
    <td>
    접수상태:  
    <select name="ForceCurrState" class="select" OnChange="jsReloadMe(this);">
    <option value="B001" <%= ChkIIF(orgCurrState="B001","selected","") %> >접수
    <option value="B007" <%= ChkIIF(orgCurrState="B007","selected","") %> >완료
    </select>  
    &nbsp;&nbsp;&nbsp;&nbsp;
    받는사람 메일:
    <input type="text" name="ForceBuyEmail" class="text" value="<%= orgBuyEmail %>" size="30" maxlength="100">
    &nbsp;
    <input type="button" class="button" value="재발송" onClick="ReSendMail()">
    </td>
</tr>
</form>
</table>
<hr>
<br>
<% response.write strHTML %>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->