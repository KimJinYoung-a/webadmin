<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim page
dim opartner, mg
dim mduserid, rect, isusing, research

page        = requestCheckVar(request("page"),10)
mduserid    = requestCheckVar(request("mduserid"),32)
rect        = requestCheckVar(request("rect"),32)
mg        = requestCheckVar(request("mg"),32)
isusing     = requestCheckVar(request("isusing"),10)
research    = requestCheckVar(request("research"),10)

if page="" then page=1
if isusing="" and research="" then isusing="on"

set opartner = new CPartnerUser
opartner.FCurrpage = page
opartner.FPageSize = 1000
opartner.FRectIsUsing = isusing
opartner.FRectInitial=rect
opartner.FRectManagerName = mg
opartner.GetAcademyPartnerList
'''opartner.GetPartnerQuickSearch

dim i
%>
<script language='javascript'>
function poplecUser(v){
    var popwin = window.open("/academy/lecture/poplecUser.asp?lecturer_id=" + v,"popupcheinfo","width=740 height=580 scrollbars=yes resizable=yes");
    popwin.focus();
}


function SearchBrand(){
//	if ((frm.mduserid.value.length<1)&&(frm.rect.value.length<2)){
//		alert('두글자 이상 입력하세요.');
//		frm.rect.focus();
//		return;
//	}

	frm.submit();
}

function PopUpcheInfo(v){
	var popwin = window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopgroupInfo(v){
	var popwin = window.open("/admin/lib/popupcheinfoonly.asp?groupid=" + v,"popupcheinfoonly","width=640 height=660 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopBrandMeachulsum(v){
	alert('준비중');
}

function ExcelPrint() {
	xlfrm.target="iiframeXL";
	xlfrm.action="/academy/lecture/dolecturerlistexcel.asp";
	xlfrm.submit();
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="rectorder" value="">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
				아이디 : <input type="text" name="rect" value="<%= rect %>" Maxlength="32" size="16"> (앞 두글자 이상 : 영문)
				&nbsp;
				담당자 : <input type="text" name="mg" value="<%= mg %>" Maxlength="32" size="16"> (앞 두글자 이상 : 한글)
				&nbsp;
				<input type="checkbox" name="isusing" <%= chkIIF(isusing="on","checked","") %> >사용강사만
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<a href="javascript:SearchBrand();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->



<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#bababa">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right">총 <%= opartner.FtotalCount %>건&nbsp;<input type="button" onclick="ExcelPrint();" value="엑셀다운로드" class="button"></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	
	<td width="70" rowspan=2>강사ID</td>
	<td width="40" rowspan=2>강좌</td>
	<td width="40" rowspan=2>마진</td>
	<td width="40" rowspan=2>재료비<br>마진</td>
	<td width="40" rowspan=2>DIY</td>
	<td width="40" rowspan=2>마진</td>
<!--	<td width="30" rowspan=2>SCM</td>  -->
	<td width="80" rowspan=2>스트리트명</td>
	<td width="80" rowspan=2>회사명</td>
	<td width="60" rowspan=2>담당자</td>
	<td width="90" rowspan=2>전화번호</td>
	<td rowspan=2>E-Mail / 등록일</td>
	<td width="70" colspan=2>사용여부</td>
	<td width="105" colspan=3>스트리트오픈여부</td>
	<td width="30" rowspan=2>매출<br>추이</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td width="35">텐바<br>이텐</td>
	<td width="35">제휴<br>몰</td>
	<td width="35">텐바<br>이텐</td>
	<td width="35">제휴<br>몰</td>
	<td width="35">커뮤<br>니티</td>
</tr>
<% for i=0 to opartner.FresultCount-1 %>
<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
<tr bgcolor="#FFFFFF">
<% else %>
<tr bgcolor="#EEEEEE">
<% end if %>
	<td><a href="javascript:PopBrandInfoEdit('<%= opartner.FPartnerList(i).FID %>')"><%= opartner.FPartnerList(i).FID %></a></td>
	<td align="center">
	    <% if IsNULL(opartner.FPartnerList(i).Flec_yn) then %>
	    <img src="/images/icon_arrow_link.gif" border="0" style="cursor:pointer" onClick="poplecUser('<%= opartner.FPartnerList(i).FID %>')">
	    <% else %>
	    <a href="javascript:poplecUser('<%= opartner.FPartnerList(i).FID %>');"><%= opartner.FPartnerList(i).Flec_yn %></a>
	    <% end if %>
	</td>
	<td align="center"><a href="javascript:poplecUser('<%= opartner.FPartnerList(i).FID %>');"><%= opartner.FPartnerList(i).Flec_margin %></a></td>
	<td align="center"><a href="javascript:poplecUser('<%= opartner.FPartnerList(i).FID %>');"><%= opartner.FPartnerList(i).Fmat_margin %></a></td>
	<td align="center">
	    <% if IsNULL(opartner.FPartnerList(i).Fdiy_yn) then %>
	    <img src="/images/icon_arrow_link.gif" border="0" style="cursor:pointer" onClick="poplecUser('<%= opartner.FPartnerList(i).FID %>')">
	    <% else %>
	    <a href="javascript:poplecUser('<%= opartner.FPartnerList(i).FID %>');"><%= opartner.FPartnerList(i).Fdiy_yn %></a>
	    <% end if %>
	</td>
	<td align="center"><a href="javascript:poplecUser('<%= opartner.FPartnerList(i).FID %>');"><%= opartner.FPartnerList(i).Fdiy_margin %></a></td>
<!--
	<% if IsNull(opartner.FPartnerList(i).Fpid) or (opartner.FPartnerList(i).Fpid="") then %>
		<td bgcolor="#FF0000" align="center">-</td>
	<% else %>
		<td align="center">O</td>
	<% end if %>
-->
	<td>
		<%= opartner.FPartnerList(i).FSocName_Kor %><br>
		<%= opartner.FPartnerList(i).FSocName %>
	</td>
	<td><a href="javascript:PopUpcheInfoEdit('<%= opartner.FPartnerList(i).FGroupID %>')"><%= opartner.FPartnerList(i).Fcompany_name %></a></td>
	<td align="center"><%= opartner.FPartnerList(i).Fmanager_name %></td>
	<td>
		<%= opartner.FPartnerList(i).Ftel %><br>
		<%= opartner.FPartnerList(i).Fmanager_hp %>
		<!-- <br>Fax:<%= opartner.FPartnerList(i).Ffax %> -->
	</td>
	<td><a target=_blank href="mailto:<%= opartner.FPartnerList(i).Femail %>"><%= opartner.FPartnerList(i).Femail %></a><br>
	<%= opartner.FPartnerList(i).Fregdate %>
	</td>
	<td align=center>
	<% if opartner.FPartnerList(i).Fisusing="Y" then %>
	O
	<% else %>
	X
	<% end if %>
	</td>
	<td align=center>
	<% if opartner.FPartnerList(i).Fisextusing="Y"	then %>
	O
	<% else %>
	X
	<% end if %>
	</td>
	<td align=center>
	<% if opartner.FPartnerList(i).Fstreetusing="Y"	then %>
	O
	<% else %>
	X
	<% end if %>
	</td>
	<td align=center>
	<% if opartner.FPartnerList(i).Fextstreetusing="Y" then %>
	O
	<% else %>
	X
	<% end if %>
	</td>
	<td align=center>
	<% if opartner.FPartnerList(i).Fspecialbrand="Y" then %>
	O
	<% else %>
	X
	<% end if %>
	</td>
	<td align=center><a href="javascript:PopBrandMeachulsum();"><img src="/images/icon_arrow_link.gif" border=0></a></td>
</tr>
<% next %>
</table>
<%
set opartner = Nothing
%>
<form name="xlfrm" method="post" action="">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="mduserid" value="<%= mduserid %>">
<input type="hidden" name="rect" value="<%= rect %>">
<input type="hidden" name="isusing" value="<%= isusing %>">
<input type="hidden" name="research" value="<%= research %>">
</form>
<iframe name="iiframeXL" name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->