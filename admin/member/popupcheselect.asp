<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  OFFSHOP 정보
' History : 2009.04.07 서동석 최초생성
'			2012.07.31 한용민 매장 공용으로 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim ogroup, frmname, page, rectconame,rectDesigner, rectsocno ,compname, i , menupos ,shopyn, groupid, mode, pcuserdiv
	frmname         = request("frmname")
	page            = requestCheckVar(request("page"),9)
	rectconame      = requestCheckVar(request("rectconame"),32)
	rectDesigner    = requestCheckVar(request("rectDesigner"),32)
	rectsocno       = requestCheckVar(request("rectsocno"),16)
	groupid         = requestCheckVar(request("groupid"),16)
	compname        = request("compname")
	menupos         = RequestCheckVar(request("menupos"),10)
	shopyn          = RequestCheckVar(request("shopyn"),1)
	mode            = requestCheckVar(request("mode"),32)
    pcuserdiv       = requestCheckVar(request("pcuserdiv"),32)

if page="" then page=1

set ogroup = new CPartnerGroup
	ogroup.FPageSize = 15
	ogroup.FCurrPage = page
	ogroup.FrectDesigner = rectDesigner
	ogroup.Frectconame = rectconame
	ogroup.FRectsocno = rectsocno
	ogroup.FRectGroupid = groupid
    ogroup.FRectPCuserDiv = pcuserdiv

	if (rectDesigner<>"") then
		ogroup.GetGroupInfoListByBrand
	else
		ogroup.GetGroupInfoList
	end if

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value=page;
    document.frm.submit();

}

<% if (compname<>"") then %>
function SelectThis(frmbuf){
	var openformcomp = eval('opener.<%= frmname %>.<%= compname %>');
	openformcomp.value = frmbuf.groupid.value;
	window.close();
}

<% else %>
	<%
	' /cscenter/taxsheet/tax_view.asp
	if mode="tax" then
	%>
		function SelectThis(frmbuf){
			var openform = eval('opener.<%= frmname %>');
			openform.socname.value = frmbuf.company_name.value;
			openform.ceoname.value = frmbuf.ceoname.value;
			openform.socno.value = frmbuf.company_no.value;
			openform.socaddr.value = frmbuf.company_address.value + ' ' + frmbuf.company_address2.value;
			openform.socstatus.value = frmbuf.company_uptae.value;
			openform.socevent.value = frmbuf.company_upjong.value;
			openform.managername.value = frmbuf.manager_name.value;
			openform.managerphone.value = frmbuf.manager_phone.value;
			openform.managermail.value = frmbuf.manager_email.value;

			window.close();
		}
	<% elseif (mode = "cogs") then %>
		function SelectThis(frmbuf) {
			var openform = eval('opener.<%= frmname %>');
			openform.groupCode.value = frmbuf.groupid.value;
			window.close();
		}
	<% else %>
		function SelectThis(frmbuf){
			var openform = eval('opener.<%= frmname %>');
			<% if mode="newbrand" then %>
				opener.viewtable();
				opener.DisableSocInfo();
			<% end if %>
			openform.groupid.value = frmbuf.groupid.value;
			openform.company_name.value = frmbuf.company_name.value;
			openform.ceoname.value = frmbuf.ceoname.value;
			openform.company_no.value = frmbuf.company_no.value;
			openform.jungsan_gubun.value = frmbuf.jungsan_gubun.value;
			openform.company_zipcode.value = frmbuf.company_zipcode.value;
			openform.company_address.value = frmbuf.company_address.value;
			openform.company_address2.value = frmbuf.company_address2.value;
			openform.company_uptae.value = frmbuf.company_uptae.value;
			openform.company_upjong.value = frmbuf.company_upjong.value;
			openform.jungsan_name.value = frmbuf.jungsan_name.value;
			openform.jungsan_email.value = frmbuf.jungsan_email.value;
			openform.jungsan_hp.value = frmbuf.jungsan_hp.value;

            if ('<%= frmname %>'=='frmbrand'){
                openform.partnerCnt.value = frmbuf.partnerCnt.value;
            }


			<% if shopyn <> "Y" then %>
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
				openform.deliver_name.value = frmbuf.deliver_name.value;
				openform.deliver_phone.value = frmbuf.deliver_phone.value;
				openform.deliver_email.value = frmbuf.deliver_email.value;
				openform.deliver_hp.value = frmbuf.deliver_hp.value;
				openform.p_return_zipcode.value = frmbuf.return_zipcode.value;
				openform.p_return_address.value = frmbuf.return_address.value;
				openform.p_return_address2.value = frmbuf.return_address2.value;
				openform.defaultsongjangdiv.value = frmbuf.defaultsongjangdiv.value;
				openform.prtidx.value = frmbuf.prtidx.value;
			<% end if %>
			window.close();
		}
	<% end if %>
<% end if %>

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="frmname" value="<%= frmname %>">
<input type="hidden" name="compname" value="<%= compname %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="shopyn" value="<%= shopyn %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    그룹코드 <input type=text name=groupid value="<%= groupid %>" size=8 maxlength=10>&nbsp;
        사업자번호 <input type=text name=rectsocno value="<%= rectsocno %>" size=12 maxlength=12>&nbsp;
        회사명 <input type=text name=rectconame value="<%= rectconame %>" size=10 maxlength=32>&nbsp;
        포함브랜드 <input type="text" name="rectDesigner" value="<%= rectDesigner %>" Maxlength="32" size="16">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>

<table width="100%" border=0 cellspacing=1 cellpadding=3  class=a bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#DDDDFF" align="center">
	<td width=100>업체코드</td>
	<td width=260>업체명</td>
	<td width=140>사업자번호</td>
	<td>진행브랜드</td>
	<td width=70>선택</td>
</tr>
<% if ogroup.FResultCount>0 then %>
<% for i=0 to ogroup.FResultCount -1 %>
<form name=frm_<%= i %> >
<input type="hidden" name=groupid value="<%= ogroup.FItemList(i).FGroupID %>">
<input type="hidden" name=company_name value="<%= ogroup.FItemList(i).Fcompany_name %>">
<input type="hidden" name=ceoname value="<%= ogroup.FItemList(i).Fceoname %>">
<input type="hidden" name=company_no value="<%= ogroup.FItemList(i).Fcompany_no %>">
<input type="hidden" name=jungsan_gubun value="<%= ogroup.FItemList(i).Fjungsan_gubun %>">
<input type="hidden" name=company_zipcode value="<%= ogroup.FItemList(i).Fcompany_zipcode %>">
<input type="hidden" name=company_address value="<%= ogroup.FItemList(i).Fcompany_address %>">
<input type="hidden" name=company_address2 value="<%= ogroup.FItemList(i).Fcompany_address2 %>">
<input type="hidden" name=company_uptae value="<%= ogroup.FItemList(i).Fcompany_uptae %>">
<input type="hidden" name=company_upjong value="<%= ogroup.FItemList(i).Fcompany_upjong %>">
<input type="hidden" name=company_tel value="<%= ogroup.FItemList(i).Fcompany_tel %>">
<input type="hidden" name=company_fax value="<%= ogroup.FItemList(i).Fcompany_fax %>">
<input type="hidden" name=jungsan_bank value="<%= ogroup.FItemList(i).Fjungsan_bank %>">
<input type="hidden" name=jungsan_acctno value="<%= ogroup.FItemList(i).Fjungsan_acctno %>">
<input type="hidden" name=jungsan_acctname value="<%= ogroup.FItemList(i).Fjungsan_acctname %>">
<input type="hidden" name=jungsan_date value="<%= ogroup.FItemList(i).Fjungsan_date %>">
<input type="hidden" name=jungsan_date_off value="<%= ogroup.FItemList(i).Fjungsan_date_off %>">
<input type="hidden" name=manager_name value="<%= ogroup.FItemList(i).Fmanager_name %>">
<input type="hidden" name=manager_phone value="<%= ogroup.FItemList(i).Fmanager_phone %>">
<input type="hidden" name=manager_email value="<%= ogroup.FItemList(i).Fmanager_email %>">
<input type="hidden" name=manager_hp value="<%= ogroup.FItemList(i).Fmanager_hp %>">
<input type="hidden" name=deliver_name value="<%= ogroup.FItemList(i).Fdeliver_name %>">
<input type="hidden" name=deliver_phone value="<%= ogroup.FItemList(i).Fdeliver_phone %>">
<input type="hidden" name=deliver_email value="<%= ogroup.FItemList(i).Fdeliver_email %>">
<input type="hidden" name=deliver_hp value="<%= ogroup.FItemList(i).Fdeliver_hp %>">
<input type="hidden" name=jungsan_name value="<%= ogroup.FItemList(i).Fjungsan_name %>">
<input type="hidden" name=jungsan_phone value="<%= ogroup.FItemList(i).Fjungsan_phone %>">
<input type="hidden" name=jungsan_email value="<%= ogroup.FItemList(i).Fjungsan_email %>">
<input type="hidden" name=jungsan_hp value="<%= ogroup.FItemList(i).Fjungsan_hp %>">
<input type="hidden" name=return_zipcode value="<%= ogroup.FItemList(i).Freturn_zipcode %>">
<input type="hidden" name=return_address value="<%= ogroup.FItemList(i).Freturn_address %>">
<input type="hidden" name=return_address2 value="<%= ogroup.FItemList(i).Freturn_address2 %>">
<input type="hidden" name=defaultsongjangdiv value="<%= ogroup.FItemList(i).Fdefaultsongjangdiv %>">
<input type="hidden" name=prtidx value="<%= ogroup.FItemList(i).FPrtIdx %>">
<input type="hidden" name=partnerCnt value="<%= ogroup.FItemList(i).fpartnerCnt %>">
<tr bgcolor="#FFFFFF">
	<td><%= ogroup.FItemList(i).FGroupID %></td>
	<td><%= ogroup.FItemList(i).FCompany_Name %></td>
	<td><%= socialnoReplace(ogroup.FItemList(i).FCompany_No) %></td>
	<td <%=ChkIIF(ogroup.FItemList(i).getPartnerIdInfoStr="","bgcolor='#CCCCCC'","")%> ><%= ogroup.FItemList(i).getPartnerIdInfoStr %></td>
	<td width=70 align="center">
		<input type="button" value="선택" onClick="SelectThis(frm_<%= i %>)" class="button">
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF" height=30>
	<td colspan=10 align=center>
	<% if ogroup.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ogroup.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ogroup.StartScrollPage to ogroup.FScrollCount + ogroup.StartScrollPage - 1 %>
		<% if i>ogroup.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ogroup.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=10 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>
</table>

<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
