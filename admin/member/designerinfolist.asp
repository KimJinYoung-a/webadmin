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
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/util/base64unicode.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim isLocalIP : isLocalIP = fn_isDongSoongIP()
dim makerid ,catecode, groupid, offcatecode, offmduserid ,mrectTp ,i,page, pcuserdiv, readypartner   ''' partner'userdiv _ user_c'useridv
dim usingonly, research, userdiv, rect, crect, mrect, mduserid, companyno, itemid,socname_kr, Stdate, Eddate, purchasetype, qstring
dim jungsan_gubun, dispCate
	pcuserdiv   = RequestCheckVar(request("pcuserdiv"),32)
	makerid     = RequestCheckVar(request("makerid"),32)
	usingonly   = request("usingonly")
	research    = request("research")
	userdiv     = RequestCheckVar(request("userdiv"),32)
	rect        = RequestCheckVar(request("rect"),32)
	socname_kr  = requestCheckVar(request("socname_kr"),60)
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
	Stdate     = RequestCheckVar(request("Stdate"),10)
	Eddate     = RequestCheckVar(request("Eddate"),10)
	purchasetype     = RequestCheckVar(request("purchasetype"),10)
	readypartner     = RequestCheckVar(request("readypartner"),2)
	jungsan_gubun     = RequestCheckVar(request("jungsan_gubun"),10)
	dispCate	= RequestCheckVar(request("dispCate"),3)

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
	opartner.FRectSOCName  = socname_kr
	opartner.FRectCompanyname = crect

	if mrectTp = "dname" then
		opartner.FRectManagerName = mrect
	elseif mrectTp = "demail" then
		opartner.FRectManageremail = mrect
	elseif mrectTp = "dphone" then
		opartner.FRectManagerhp = mrect
	end if
	if jungsan_gubun<>"" then
		opartner.FRectJungsanGubun = jungsan_gubun
	end if
	opartner.FRectCatecode = catecode
	opartner.Fitemid = itemid
	opartner.FRectCompanyNo = replace(companyno,"-","")
	opartner.FRectoffcatecode = offcatecode
	opartner.FRectoffmduserid = offmduserid
	opartner.FRectStdate = Stdate
	opartner.FRectEddate = Eddate
	opartner.FRectpurchasetype = purchasetype
	opartner.FRectReadyPartner = readypartner
	opartner.FRectDispCate = dispCate
	opartner.GetPartnerSearch
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="JavaScript" src="/js/jquery-1.7.2.min.js"></script>
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
	var popwin = window.open("/admin/member/addnewbrand.asp","addnewbrand","width=1400 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function AddNewBrand2(){
	var popwin = window.open("/admin/member/addnewbrand_step1.asp","addnewbrand2","width=1400 height=800 scrollbars=yes resizable=yes");
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

function AddNewUpcheReg(qs){
	var popwin = window.open("/common/partner/companyinfo.asp?qs="+qs,"addnewbrand2","width=1200 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
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
    for (var i=0; i<frm.itemid.value.length; i++){
        chr1 = frm.itemid.value.charAt(i);
        if(!(chr1 >= '0' && chr1 <= '9')) {
            alert("��ǰ��ȣ�� ���ڸ� �Է��ϼ���.");
            frm.itemid.focus();
            return false;
        }
    }

	if (frm.Stdate.value != "") {
		if (frm.Stdate.value.length != 10) {
			alert('�߸��� ��¥�Դϴ�.');
            frm.Stdate.focus();
            return false;
		}
	}

	if (frm.Eddate.value != "") {
		if (frm.Eddate.value.length != 10) {
			alert('�߸��� ��¥�Դϴ�.');
            frm.Eddate.focus();
            return false;
		}
	}

	frm.submit();
}

function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popShopInfo(ishopid){
	var popwin = window.open("/admin/lib/popoffshopinfo.asp?shopid=" + ishopid + "&menupos=277","popoffshopinfo",'width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popSendJoinInfo(brandid,qs){
	var popwin = window.open("/admin/member/reSendJoinInfo.asp?brandid=" + brandid + "&qs=" + qs,"popjoinpage",'width=500,height=300,scrollbars=yes,resizable=yes');
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
    <td rowspan="3" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
    <td align="left">
    	<input type="radio" name="pcuserdiv" value="" <% if pcuserdiv="" then response.write "checked" %> >��ü
        <input type="radio" name="pcuserdiv" value="9999_02" <% if pcuserdiv="9999_02" then response.write "checked" %> >����ó(�Ϲ�)
        <input type="radio" name="pcuserdiv" value="9999_14" <% if pcuserdiv="9999_14" then response.write "checked" %> >����ó(����)
        <% if (FALSE) then %>
        <input type="radio" name="pcuserdiv" value="9999_15" <% if pcuserdiv="9999_15" then response.write "checked" %> >����ó(Fingers)
        <% end if %>
        &nbsp;|&nbsp;
        <input type="radio" name="pcuserdiv" value="999_50"  <% if pcuserdiv="999_50" then response.write "checked" %> >���޻�(�¶���)
        <input type="radio" name="pcuserdiv" value="501_21"  <% if pcuserdiv="501_21" then response.write "checked" %> >������
		<input type="radio" name="pcuserdiv" value="502_21"  <% if pcuserdiv="502_21" then response.write "checked" %> >������
        <input type="radio" name="pcuserdiv" value="503_21"  <% if pcuserdiv="503_21" then response.write "checked" %> >����ó
        <input type="radio" name="pcuserdiv" value="900_21"  <% if pcuserdiv="900_21" then response.write "checked" %> >���ó(��Ÿ)
		<input type="radio" name="pcuserdiv" value="901_21"  <% if pcuserdiv="901_21" then response.write "checked" %> >����������
		<input type="radio" name="pcuserdiv" value="902_21"  <% if pcuserdiv="902_21" then response.write "checked" %> >���¾�ü
		<input type="radio" name="pcuserdiv" value="903_21"  <% if pcuserdiv="903_21" then response.write "checked" %> >3PL(��ǥ)
        &nbsp;&nbsp;&nbsp;
        <input type="checkbox" name="usingonly" value="on" <%= CHKIIF(usingonly="on","checked","") %> > ���귣�常 ����
		&nbsp;&nbsp;&nbsp;
		<input type="checkbox" name="readypartner" value="on" <%= CHKIIF(readypartner="on","checked","") %> > ���� �������� ��ü�� ����
	</td>
	<td rowspan="3" width="50" bgcolor="#EEEEEE"><input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td align="left" >
		����ī�װ�: <%= fnStandardDispCateSelectBox(1,"", "dispCate", dispCate, "")%>
		&nbsp;
		ī�װ�ON : <% SelectBoxBrandCategory "catecode", catecode %>
		&nbsp;
		ī�װ�OFF : <% SelectBoxBrandCategory "offcatecode", offcatecode %>
		&nbsp;
		�����ON : <% drawSelectBoxCoWorker_OnOff "mduserid", mduserid, "on" %>
		&nbsp;
		�����OFF : <% drawSelectBoxCoWorker_OnOff "offmduserid", offmduserid, "off" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" >
    <td align="left" >
        �귣��ID <input type="text" name="rect" value="<%= rect %>" Maxlength="32" size="14">
        &nbsp;
		�׷��ڵ� <input type="text" name="groupid" value="<%= groupid %>" Maxlength="32" size="7">
		&nbsp;
		��Ʈ��Ʈ��(�ѱ�) : <input type="text" class="text" name="socname_kr" value="<%= socname_kr %>" Maxlength="32" size="16">
		&nbsp;
		ȸ��� <input type="text" name="crect" value="<%= crect %>" Maxlength="32" size="12">
		&nbsp;
		����ڹ�ȣ <input type="text" name="companyno" value="<%=companyno %>" Maxlength="32" size="12">
		<br>
		<select name="mrectTp">
			<option value="dname"  <%=CHKIIF(mrectTp="dname","selected","") %> >����ڸ�
			<option value="demail" <%=CHKIIF(mrectTp="demail","selected","") %> >�����Email
			<option value="dphone" <%=CHKIIF(mrectTp="dphone","selected","") %> >����ڿ���ó
		</select>
		<input type="text" name="mrect" value="<%= mrect %>" Maxlength="32" size="10">
		&nbsp;
		��ǰ��ȣ <input type="text" name="itemid" value="<%=itemid%>" size="8" />
		&nbsp;
		����� :
		<input id="Stdate" name="Stdate" value="<%=Stdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="Stdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="Eddate" name="Eddate" value="<%=Eddate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="Eddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "Stdate", trigger    : "Stdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "Eddate", trigger    : "Eddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;
		�������� :
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchaseType,"" %>
		&nbsp;
		���� ���� :
		<select name="jungsan_gubun" class="select" onchange="fnJungsanGubunChanged()">
			<option value="" <% if jungsan_gubun="" then response.write "selected" %>>��ü</option>
			<option value="�Ϲݰ���" <% if jungsan_gubun="�Ϲݰ���" then response.write "selected" %>>�Ϲݰ���</option>
			<option value="���̰���" <% if jungsan_gubun="���̰���" then response.write "selected" %>>���̰���</option>
			<option value="��õ¡��" <% if jungsan_gubun="��õ¡��" then response.write "selected" %>>��õ¡��</option>
			<option value="�鼼" <% if jungsan_gubun="�鼼" then response.write "selected" %>>�鼼</option>
			<option value="����(�ؿ�)" <% if jungsan_gubun="����(�ؿ�)" then response.write "selected" %>>����(�ؿ�)</option>
		</select>
    </td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" class="a" >
<tr>
	<td align="left">

	</td>
	<td align="right">
		<input type=button value="��������ȭ���" onclick="AddNewBrand2();" class="button"> <input type=button value="�űԾ�ü���" onclick="AddNewBrand();" class="button">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= opartner.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= opartner.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan=2>����</td>
	<td rowspan=2>�귣��ID</td>
	<td rowspan=2>�귣���(�ѱ�)<br>�귣���(����)</td>
	<td rowspan=2>�׷��ڵ�<br>����ڹ�ȣ</td>
	<td rowspan=2>ȸ���</td>
	<td rowspan=2>��������</td>
	<td rowspan=2>�����</td>
	<td rowspan=2>�����</td>
	<td width="90" rowspan=2>��ȭ��ȣ<br>�ڵ�����ȣ</td>
	<td width="40" rowspan=2>�̸���</td>
	<td width="70" colspan=3>��뿩��</td>
	<td rowspan=2>��ü����<br>���¿���</td>
	<td rowspan=2>�귣��<br>�߰�����</td>
	<td rowspan=2>��Ÿ����</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="35">�ٹ�����<br>ON</td>
	<td width="35">�ٹ�����<br>OFF</td>
	<td width="35">���޸�</td>
</tr>
<% if opartner.FresultCount > 0 then %>
<% for i=0 to opartner.FresultCount-1 %>
<% qstring = Server.UrlEncode(TBTEncryptUrl(Cstr(opartner.FPartnerList(i).FID) + "|" +Cstr(opartner.FPartnerList(i).FpcUserDiv))) %>
<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
<tr bgcolor="#FFFFFF">
<% else %>
<tr bgcolor="#EEEEEE">
<% end if %>
	<td align="center"><%= opartner.FPartnerList(i).GetUserDivName %></a></td>
	<td><a href="<%=vwwwUrl%>/street/street_brand.asp?makerid=<%= opartner.FPartnerList(i).FID %>" title="�귣�� ��Ʈ��Ʈ ����" target="_blank"><%= opartner.FPartnerList(i).FID %></a></td>
	<td>
		<a href="javascript:PopBrandInfoEdit('<%= opartner.FPartnerList(i).FID %>')" title="�귣�� ���� ����">
			<%= opartner.FPartnerList(i).FSocName_Kor %><br>
			<%= opartner.FPartnerList(i).FSocName %>
		</a>
	</td>
	<td <%= CHKIIF((Trim(opartner.FPartnerList(i).FGroupId)="") or isNULL(opartner.FPartnerList(i).FGroupId),"bgcolor='#EEEEEE'","") %> >
		<% if opartner.FPartnerList(i).FGroupId="" or IsNull(opartner.FPartnerList(i).FGroupId) then %>
		<a href="javascript:AddNewUpcheReg('<%=qstring%>')"><font color="red">�׷��ڵ� ����</font> </a><br>
        <% else %>
        <%= opartner.FPartnerList(i).FGroupId %><br>
		<% end if %>
		<%= socialnoBlank(opartner.FPartnerList(i).Fcompany_no) %>
	</td>
	<td><a href="javascript:PopUpcheInfoEdit('<%= opartner.FPartnerList(i).FGroupID %>')"><%= opartner.FPartnerList(i).Fcompany_name %></a></td>
	<td align="center">
		<%= opartner.FPartnerList(i).fpurchasetypename %>
	</td>
	<td align="center"><%= Left(opartner.FPartnerList(i).Fregdate,10) %></td>
	<td align="center"><%= opartner.FPartnerList(i).Fmanager_name %></td>
	<td>
	    <% if (isLocalIP) then '' md�� ��û(���) //2016/02/26
	    %>
	        <%= (opartner.FPartnerList(i).Ftel) %><br>
		    <%= (opartner.FPartnerList(i).Fmanager_hp) %>
	    <% else %>
		    <%= GetTelWithAsterisk(opartner.FPartnerList(i).Ftel) %><br>
		    <%= GetTelWithAsterisk(opartner.FPartnerList(i).Fmanager_hp) %>
	    <% end if %>
	</td>
	<td align="center">
	     <% if (isLocalIP) then %>
	<%= opartner.FPartnerList(i).Femail %>
	<%end if%>
		<% if opartner.FPartnerList(i).Femail<>"" then %>
		&nbsp;<a href="mailto:<%= opartner.FPartnerList(i).Femail %>"><img src="/images/icon_search.jpg" width="16" border="0" alt="<%= opartner.FPartnerList(i).Femail %>"></a>
		<% else %>
		&nbsp;
		<% end if %>
	</td>
	<td align=center>
		<% if opartner.FPartnerList(i).Fisusing="Y" then %>
		O
		<% else %>
		X
		<% end if %>
	</td>
	<td align=center>
		<% if opartner.FPartnerList(i).Fisoffusing="Y"	then %>
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
		<a href="javascript:PopBrandAdminUsingChange('<%= opartner.FPartnerList(i).FID %>');">
		<% if opartner.FPartnerList(i).Fpartnerusing="Y" then %>
			<% if opartner.FPartnerList(i).Fisusing="N" then %>
			<font color="red"><b>O</b></font>
			<% else %>
			O
			<% end if %>
		<% elseif IsNULL(opartner.FPartnerList(i).Fpartnerusing) then %>
		<font color="red">����</font>
		<% else %>
		<font color="red">X</font>
		<% end if %>
		</a>
	</td>
	<td align=center>
	<% if (opartner.FPartnerList(i).isbuyingPartner) then %>
	<a href="javascript:popSimpleBrandInfo('<%= opartner.FPartnerList(i).FID %>')">[����]</a>
	<% elseif (opartner.FPartnerList(i).isShopPartner) then %>
	<a href="javascript:popShopInfo('<%= opartner.FPartnerList(i).FID %>')">[����]</a>
	<% end if %>
	</td>
	<td>
		<% if opartner.FPartnerList(i).FGroupId="" or IsNull(opartner.FPartnerList(i).FGroupId) then %>
		<a href="javascript:AddNewUpcheReg('<%=qstring%>')">������������</a><br>
		<a href="javascript:popSendJoinInfo('<%= opartner.FPartnerList(i).FID %>','<%=qstring%>')">�������� ��߼�</a>
		<% end if %>
	</td>
</tr>
<% next %>

<tr bgcolor="FFFFFF">
	<td colspan="16" align="center">
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
	<td colspan="16">
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

function ereg(strOriginalString, strPattern, varIgnoreCase)
    ' Function matches pattern, returns true or false
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    ereg = objRegExp.test(strOriginalString)
    set objRegExp = nothing
end Function

function ereg_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
    ' Function replaces pattern with replacement
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    ereg_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function

''TODO : ���̳ʽ� ���� ��ȭ��ȣ ó�� ����.
''(0101112222, 021112222, 0312223333)
function GetTelWithAsterisk(telNo)
	dim resultStr, tmpArr, i

	resultStr = telNo

	if IsNull(telno) then
		GetTelWithAsterisk = resultStr
		Exit Function
	end if

	tmpArr = Split(telNo, "-")

	Select Case UBound(tmpArr)
		Case 1
			resultStr = ereg_replace(tmpArr(0), ".", "*", True) & "-" & tmpArr(0)
		Case 2
			resultStr = tmpArr(0) & "-" & ereg_replace(tmpArr(1), ".", "*", True) & "-" & tmpArr(2)
		Case Else
			resultStr = "ERR"
	End Select

	GetTelWithAsterisk = resultStr
end Function
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
