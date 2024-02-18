<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드 리스트
' History : 2012.08.21 서동석 생성
'			2012.08.22 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

dim frmName, compName
frmName= request("frmName")
compName= request("compName")

dim makerid ,catecode, groupid, offcatecode, offmduserid ,mrectTp ,i,page, pcuserdiv   ''' partner'userdiv _ user_c'useridv
dim usingonly, research, userdiv, rect, crect, mrect, mduserid, companyno, itemid
dim rectname

	pcuserdiv   = RequestCheckVar(request("pcuserdiv"),32)
	makerid     = RequestCheckVar(request("makerid"),32)
	usingonly   = request("usingonly")
	research    = request("research")
	userdiv     = RequestCheckVar(request("userdiv"),32)
	rect        = RequestCheckVar(request("rect"),32)
	mduserid    = RequestCheckVar(request("mduserid"),32)
	catecode    = RequestCheckVar(request("catecode"),32)
	crect       = RequestCheckVar(request("crect"),32)
	mrect       = RequestCheckVar(request("mrect"),64)
	companyno   = RequestCheckVar(request("companyno"),32)
	itemid		= RequestCheckVar(request("itemid"),32)
	groupid     = RequestCheckVar(request("groupid"),32)
	offcatecode = RequestCheckVar(request("offcatecode"),32)
	offmduserid = RequestCheckVar(request("offmduserid"),32)
	mrectTp     = RequestCheckVar(request("mrectTp"),32)
	rectname    = RequestCheckVar(request("rectname"),32)
	page        = request("page")

'####### 20110905 최맑은소리 사용중인것을 전체로 바꿔달라고 함.
'''if ((research="") and (usingonly="")) then usingonly="all" ''디폴트 빈값.
if page="" then page=1

dim opartner
set opartner = new CPartnerUser
	opartner.FCurrpage = page
	opartner.FPageSize = 50
	opartner.FRectPCuserDiv = pcuserdiv
	opartner.FRectGroupid = groupid
	opartner.FRectDesignerID = makerid
	opartner.FrectIsUsing = usingonly
	opartner.FRectDesignerDiv = userdiv
	opartner.FRectMdUserID = mduserid

	opartner.FRectInitial = Replace(rect,"'","''")
	opartner.FRectSOCName = rectname

	opartner.FRectCompanyname = crect

	if mrectTp = "dname" then
		opartner.FRectManagerName = mrect
	elseif mrectTp = "demail" then
		opartner.FRectManageremail = mrect
	elseif mrectTp = "dphone" then
		opartner.FRectManagerhp = mrect
	end if

	opartner.FRectCatecode = catecode
	opartner.Fitemid = itemid
	opartner.FRectCompanyNo = replace(companyno,"-","")
	opartner.FRectoffcatecode = offcatecode
	opartner.FRectoffmduserid = offmduserid
	opartner.GetPartnerSearch
%>

<script language='javascript'>

function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

function research(frm,order){
	frm.rectorder.value = order;
	frm.submit();
}

function AddNewBrand(){
	var popwin = window.open("/admin/member/addnewbrand.asp","addnewbrand","width=800 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function AnCheckNSongjangView(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 내역이 없습니다.');
		return;
	}

	var ret = confirm('선택 내역으로 품절확인 및 SMS발송을 하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idarr.value = upfrm.idarr.value + frm.id.value + ",";
				}
			}
		}
		//alert(upfrm.idarr.value);
		upfrm.submit();
	}
}
function onlyNumberInput()
{
	var code = window.event.keyCode;
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) {
		window.event.returnValue = true;
		return;
	}
	window.event.returnValue = false;
}

function checkform(frm)
{
	frm.submit();
}

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popShopInfo(ishopid){
	var popwin = window.open("/admin/lib/popoffshopinfo.asp?shopid=" + ishopid + "&menupos=277","popoffshopinfo",'width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function selectThis(selval){
    opener.<%= frmName %>.<%= compName %>.value = selval;
    window.close();
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" method="get" action="" onSubmit="return checkform(this);">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="rectorder" value="">
<input type="hidden" name="frmName" value="<%= frmName %>">
<input type="hidden" name="compName" value="<%= compName %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="3" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
    <td align="left">
    	<input type="radio" name="pcuserdiv" value="" <% if pcuserdiv="" then response.write "checked" %> >전체
        <input type="radio" name="pcuserdiv" value="9999_02" <% if pcuserdiv="9999_02" then response.write "checked" %> >매입처(일반)
        <input type="radio" name="pcuserdiv" value="9999_14" <% if pcuserdiv="9999_14" then response.write "checked" %> >매입처(아카데미)
        &nbsp;|&nbsp;
        <input type="radio" name="pcuserdiv" value="999_50"  <% if pcuserdiv="999_50" then response.write "checked" %> >제휴사(온라인)
        <input type="radio" name="pcuserdiv" value="501_21"  <% if pcuserdiv="501_21" then response.write "checked" %> >직영점
        <input type="radio" name="pcuserdiv" value="503_21"  <% if pcuserdiv="503_21" then response.write "checked" %> >기타매장
        <input type="radio" name="pcuserdiv" value="900_21"  <% if pcuserdiv="900_21" then response.write "checked" %> >출고처(기타)
	</td>
	<td rowspan="3" width="50" bgcolor="#EEEEEE"><input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td align="left" >
		<input type="checkbox" name="usingonly" value="on" <%= CHKIIF(usingonly="on","checked","") %> > 사용브랜드만 보기
    </td>
</tr>
<tr bgcolor="#FFFFFF" >
    <td align="left" >
        브랜드ID <input type="text" class="text" name="rect" value="<%= rect %>" Maxlength="32" size="14">
        &nbsp;
		브랜드명 <input type="text" class="text" name="rectname" value="<%= rectname %>" Maxlength="32" size="14">
        &nbsp;
		그룹코드 <input type="text" class="text" name="groupid" value="<%= groupid %>" Maxlength="32" size="7">
		&nbsp;
		회사명 <input type="text" class="text" name="crect" value="<%= crect %>" Maxlength="32" size="12">
    </td>
</tr>
</table>
<!-- 표 상단바 끝-->

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%= opartner.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= opartner.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>구분</td>
	<td>브랜드ID</td>
	<td>브랜드명(한글)<br>브랜드명(영문)</td>
	<td>그룹코드<br>사업자번호</td>
	<td>회사명</td>
</tr>
<% if opartner.FresultCount > 0 then %>
<% for i=0 to opartner.FresultCount-1 %>
<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
<tr bgcolor="#FFFFFF">
<% else %>
<tr bgcolor="#EEEEEE">
<% end if %>
	<td align="center"><%= opartner.FPartnerList(i).GetUserDivName %></a></td>
	<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>')"><%= opartner.FPartnerList(i).FID %></a></td>
	<td>
			<%= opartner.FPartnerList(i).FSocName_Kor %><br>
			<%= opartner.FPartnerList(i).FSocName %>
	</td>
	<td <%= CHKIIF((Trim(opartner.FPartnerList(i).FGroupId)="") or isNULL(opartner.FPartnerList(i).FGroupId),"bgcolor='#EEEEEE'","") %> >
		<%= opartner.FPartnerList(i).FGroupId %><br>
		<%= opartner.FPartnerList(i).Fcompany_no %>
	</td>
	<td><%= opartner.FPartnerList(i).Fcompany_name %></td>
</tr>
<% next %>

<tr bgcolor="FFFFFF">
	<td colspan="5" align="center">
    	<% if opartner.HasPreScroll then %>
		<a href="javascript:NextPage('<%= opartner.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + opartner.StartScrollPage to opartner.FScrollCount + opartner.StartScrollPage - 1 %>
			<% if i>opartner.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if opartner.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=5>
		검색 결과가 없습니다.
	</td>
</tr>
<% end if %>
</table>

<%
set opartner = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->