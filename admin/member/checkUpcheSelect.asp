<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  OFFSHOP 정보
' History : 2009.04.07 서동석 생성
'			2012.07.31 한용민 매장 공용으로 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim ogroup, frmname, page, rectconame,rectDesigner, rectsocno ,compname, i , menupos ,shopyn, groupid, mode
dim businessgubun
	frmname         = request("frmname")
	rectsocno       = requestCheckVar(request("company_no"),16)
	groupid         = requestCheckVar(request("groupid"),16)
	mode      = requestCheckVar(request("mode"),32)
    businessgubun      = requestCheckVar(request("businessgubun"),1)
	
if page="" then page=1

set ogroup = new CPartnerGroup
	ogroup.FPageSize = 15
	ogroup.FCurrPage = page
	ogroup.FrectDesigner = rectDesigner
	ogroup.Frectconame = rectconame
	ogroup.FRectsocno = rectsocno
	ogroup.FRectGroupid = groupid
	
	if (rectsocno<>"") then
		ogroup.GetGroupInfoList
	end if

%>
<% if ogroup.FResultCount>0 then %>
<form name="ufrm">
<input type="hidden" name="groupid" value="<%= ogroup.FItemList(0).FGroupID %>">
<input type="hidden" name="company_name" value="<%= ogroup.FItemList(0).Fcompany_name %>">
<input type="hidden" name="ceoname" value="<%= ogroup.FItemList(0).Fceoname %>">
<input type="hidden" name="company_no" value="<%= ogroup.FItemList(0).Fcompany_no %>">
<input type="hidden" name="jungsan_gubun" value="<%= ogroup.FItemList(0).Fjungsan_gubun %>">
<input type="hidden" name="company_zipcode" value="<%= ogroup.FItemList(0).Fcompany_zipcode %>">
<input type="hidden" name="company_address" value="<%= ogroup.FItemList(0).Fcompany_address %>">
<input type="hidden" name="company_address2" value="<%= ogroup.FItemList(0).Fcompany_address2 %>">
<input type="hidden" name="company_uptae" value="<%= ogroup.FItemList(0).Fcompany_uptae %>">
<input type="hidden" name="company_upjong" value="<%= ogroup.FItemList(0).Fcompany_upjong %>">
<input type="hidden" name="company_tel" value="<%= ogroup.FItemList(0).Fcompany_tel %>">
<input type="hidden" name="company_fax" value="<%= ogroup.FItemList(0).Fcompany_fax %>">
<input type="hidden" name="jungsan_bank" value="<%= ogroup.FItemList(0).Fjungsan_bank %>">
<input type="hidden" name="jungsan_acctno" value="<%= ogroup.FItemList(0).Fjungsan_acctno %>">
<input type="hidden" name="jungsan_acctname" value="<%= ogroup.FItemList(0).Fjungsan_acctname %>">
<input type="hidden" name="jungsan_date" value="<%= ogroup.FItemList(0).Fjungsan_date %>">
<input type="hidden" name="jungsan_date_off" value="<%= ogroup.FItemList(0).Fjungsan_date_off %>">
<input type="hidden" name="manager_name" value="<%= ogroup.FItemList(0).Fmanager_name %>">
<input type="hidden" name="manager_phone" value="<%= ogroup.FItemList(0).Fmanager_phone %>">
<input type="hidden" name="manager_email" value="<%= ogroup.FItemList(0).Fmanager_email %>">
<input type="hidden" name="manager_hp" value="<%= ogroup.FItemList(0).Fmanager_hp %>">
<input type="hidden" name="deliver_name" value="<%= ogroup.FItemList(0).Fdeliver_name %>">
<input type="hidden" name="deliver_phone" value="<%= ogroup.FItemList(0).Fdeliver_phone %>">
<input type="hidden" name="deliver_email" value="<%= ogroup.FItemList(0).Fdeliver_email %>">
<input type="hidden" name="deliver_hp" value="<%= ogroup.FItemList(0).Fdeliver_hp %>">
<input type="hidden" name="jungsan_name" value="<%= ogroup.FItemList(0).Fjungsan_name %>">
<input type="hidden" name="jungsan_phone" value="<%= ogroup.FItemList(0).Fjungsan_phone %>">
<input type="hidden" name="jungsan_email" value="<%= ogroup.FItemList(0).Fjungsan_email %>">
<input type="hidden" name="jungsan_hp" value="<%= ogroup.FItemList(0).Fjungsan_hp %>">
<input type="hidden" name="return_zipcode" value="<%= ogroup.FItemList(0).Freturn_zipcode %>">
<input type="hidden" name="return_address" value="<%= ogroup.FItemList(0).Freturn_address %>">
<input type="hidden" name="return_address2" value="<%= ogroup.FItemList(0).Freturn_address2 %>">
<input type="hidden" name="defaultsongjangdiv" value="<%= ogroup.FItemList(0).Fdefaultsongjangdiv %>">
<input type="hidden" name="prtidx" value="<%= ogroup.FItemList(0).FPrtIdx %>">
<input type="hidden" name="partnerCnt" value="<%= ogroup.FItemList(0).fpartnerCnt %>">
</form>

<script type='text/javascript'>
function SelectThis(){
    var frmbuf = document.ufrm;
    var openform = eval('parent.document.frmbrand');
    parent.viewtable();
    parent.DisableSocInfo();
    openform.groupid.value = frmbuf.groupid.value;
    openform.company_name.value = frmbuf.company_name.value;
    openform.ceoname.value = frmbuf.ceoname.value;
    openform.company_no.value = frmbuf.company_no.value;
    <%
    ' 해외사업자
    if businessgubun="5" then
    %>
        openform.company_no3.value = frmbuf.company_no.value;
    <%
    ' 원천징수
    elseif businessgubun="3" then
    %>
        openform.company_no2.value = frmbuf.company_no.value;
    <%
    ' 일반(간이)사업자
    elseif cstr(businessgubun)=cstr("1") then
    %>
        openform.company_no1.value = frmbuf.company_no.value;
    <% end if %>
    openform.jungsan_gubun.value = frmbuf.jungsan_gubun.value;
    openform.company_zipcode.value = frmbuf.company_zipcode.value;
    openform.company_address.value = frmbuf.company_address.value;
    openform.company_address2.value = frmbuf.company_address2.value;
    openform.company_uptae.value = frmbuf.company_uptae.value;
    openform.company_upjong.value = frmbuf.company_upjong.value;
    openform.jungsan_name.value = frmbuf.jungsan_name.value;
    openform.jungsan_email.value = frmbuf.jungsan_email.value;
    openform.jungsan_hp.value = frmbuf.jungsan_hp.value;	
    openform.return_zipcode.value = frmbuf.return_zipcode.value;	
    openform.return_address.value = frmbuf.return_address.value;
    openform.return_address2.value = frmbuf.return_address2.value;
    openform.partnerCnt.value = frmbuf.partnerCnt.value;
    openform.jungsan_phone.value = frmbuf.jungsan_phone.value;
    openform.company_tel.value = frmbuf.company_tel.value;
    openform.company_fax.value = frmbuf.company_fax.value;
    openform.jungsan_bank.value = frmbuf.jungsan_bank.value;
    openform.jungsan_acctno.value = frmbuf.jungsan_acctno.value;
    openform.jungsan_acctname.value = frmbuf.jungsan_acctname.value;
    openform.jungsan_date.value = frmbuf.jungsan_date.value;
    openform.jungsan_date_off.value = frmbuf.jungsan_date_off.value;
    openform.manager_name.value = frmbuf.manager_name.value;
    openform.manager_phone.value = frmbuf.manager_phone.value;
    openform.manager_email.value = frmbuf.manager_email.value;
    openform.manager_hp.value = frmbuf.manager_hp.value;
    openform.defaultsongjangdiv.value = frmbuf.defaultsongjangdiv.value;
    openform.partcheck.value="Y";
    openform.mode.value="addnewupchebrand2";
}
SelectThis();
</script>
<%
set ogroup = Nothing
%>
<% else %>
<script language="JavaScript" src="/js/jquery-1.7.2.min.js"></script>
<script type='text/javascript'>
var openform = eval('parent.document.frmbrand');
openform.partcheck.value="Y";
$("#bizCheck",parent.document).html("신규업체 대상입니다. 입점을 계속 진행해 주세요.");
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->