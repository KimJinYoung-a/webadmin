<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<%
dim id, ForceCurrState, ForceBuyEmail, resend
id              = RequestCheckVar(request("id"),10)
ForceCurrState  = RequestCheckVar(request("ForceCurrState"),10)
resend          = RequestCheckVar(request("resend"),10)

dim oCsAction,strHTML, orgBuyEmail, orgCurrState
Set oCsAction = New CsActionMailCls
if (ForceCurrState<>"") then oCsAction.FRectForceCurrState = ForceCurrState
if (ForceBuyEmail<>"") then oCsAction.FRectForceBuyEmail = ForceBuyEmail

strHTML = oCsAction.makeMailTemplate(id)

orgBuyEmail     = oCsAction.FBuyEmail
orgCurrState    = oCsAction.FCurrState

if (ForceBuyEmail<>"") then orgBuyEmail = ForceBuyEmail
if (ForceCurrState<>"") then orgCurrState = ForceCurrState


if (resend="on") then
    call ReSendCsActionMail(id, ForceCurrState, ForceBuyEmail)
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
<input type="hidden" name="id" value="<%= id %>">
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
<p>
<% response.write strHTML %>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->