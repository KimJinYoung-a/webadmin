<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �귣�� ����Ʈ
' History : 2012.08.21 ������ ����
'			2012.08.22 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim makerid ,catecode, groupid, offcatecode, offmduserid ,mrectTp ,i,page, pcuserdiv   ''' partner'userdiv _ user_c'useridv
dim usingonly, research, userdiv, rect, crect, mrect, mduserid, companyno, itemid
	pcuserdiv   = "999_50"
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
	page        = request("page")

'####### 20110905 �ָ����Ҹ� ������ΰ��� ��ü�� �ٲ�޶�� ��.
'''if ((research="") and (usingonly="")) then usingonly="all" ''����Ʈ ��.
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
		alert('���� ������ �����ϴ�.');
		return;
	}

	var ret = confirm('���� �������� ǰ��Ȯ�� �� SMS�߼��� �Ͻðڽ��ϱ�?');
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
    var chr1;

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
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" method="get" action="" onSubmit="return checkform(this);">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="rectorder" value="">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
    <td align="left">
        �귣��ID <input type="text" class="text" name="rect" value="<%= rect %>" Maxlength="32" size="14">
        &nbsp;
		�׷��ڵ� <input type="text" class="text" name="groupid" value="<%= groupid %>" Maxlength="32" size="7">
		&nbsp;
		ȸ��� <input type="text" class="text" name="crect" value="<%= crect %>" Maxlength="32" size="12">
		&nbsp;
		����ڹ�ȣ <input type="text" class="text" name="companyno" value="<%=companyno %>" Maxlength="32" size="12">
		&nbsp;
		<select name="mrectTp" class="select">
			<option value="dname"  <%=CHKIIF(mrectTp="dname","selected","") %> >����ڸ�
			<option value="demail" <%=CHKIIF(mrectTp="demail","selected","") %> >�����Email
			<option value="dphone" <%=CHKIIF(mrectTp="dphone","selected","") %> >����ڿ���ó
		</select>
		<input type="text" class="text" name="mrect" value="<%= mrect %>" Maxlength="32" size="10">
	</td>
	<td rowspan="2" width="50" bgcolor="#EEEEEE"><input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td align="left" >
		<input type="checkbox" name="usingonly" value="on" <%= CHKIIF(usingonly="on","checked","") %> > ���귣�常 ����
    </td>
</tr>
</form>
</table>
<!-- ǥ ��ܹ� ��-->

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= opartner.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= opartner.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>�귣��ID</td>
	<td>�귣���(�ѱ�)<br>�귣���(����)</td>
	<!--
	<td>�׷��ڵ�<br>����ڹ�ȣ</td>
	<td>ȸ���</td>
	-->
	<td>����μ�</td>
	<td>�������</td>
	<td>������</td>
	<td>������</td>
	<td>��꼭����</td>
	<td>��������</td>
	<td>�������</td>
	<!-- <td>��Ÿ����</td> -->
</tr>
<% if opartner.FresultCount > 0 then %>
<% for i=0 to opartner.FresultCount-1 %>
<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
<tr bgcolor="#FFFFFF">
<% else %>
<tr bgcolor="#EEEEEE">
<% end if %>
	<td align="center"><%= opartner.FPartnerList(i).GetUserDivName %></a></td>
	<td><a href="javascript:PopBrandInfoEdit('<%= opartner.FPartnerList(i).FID %>')"><%= opartner.FPartnerList(i).FID %></a></td>
	<td>
		<a href="javascript:PopBrandInfoEdit('<%= opartner.FPartnerList(i).FID %>')">
			<%= opartner.FPartnerList(i).FSocName_Kor %><br>
			<%= opartner.FPartnerList(i).FSocName %>
		</a>
	</td>
	<!--
	<td <%= CHKIIF((Trim(opartner.FPartnerList(i).FGroupId)="") or isNULL(opartner.FPartnerList(i).FGroupId),"bgcolor='#EEEEEE'","") %> >
		<%= opartner.FPartnerList(i).FGroupId %><br>
		<%= opartner.FPartnerList(i).Fcompany_no %>
	</td>
	<td><a href="javascript:PopUpcheInfoEdit('<%= opartner.FPartnerList(i).FGroupID %>')"><%= opartner.FPartnerList(i).Fcompany_name %></a></td>
	-->
	<td><%= opartner.FPartnerList(i).FsellBizNm %></td>
	<td><%= opartner.FPartnerList(i).FselltypeNm %></td>
	<td><%= opartner.FPartnerList(i).FpurchasetypeNm %></td>
	<td align="right"><%= FormatNumber((opartner.FPartnerList(i).Fcommission*100), 0) %>% &nbsp;&nbsp;</td>
	<td><%= opartner.FPartnerList(i).FtaxevaltypeNm %></td>
	<td align="center"><%= opartner.FPartnerList(i).FpmallSellTypeNm %></td>
	<td align="center"><%= opartner.FPartnerList(i).FpcomTypeNm %></td>
</tr>
<% next %>

<tr bgcolor="FFFFFF">
	<td colspan="15" align="center">
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
	<td colspan=15>
		�˻� ����� �����ϴ�.
	</td>
</tr>
<% end if %>
</table>

<form name="frmArrupdate" method="post" action="soldout_comparison_ok.asp">
	<input type="hidden" name="idarr" value="">
</form>
<%
set opartner = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->