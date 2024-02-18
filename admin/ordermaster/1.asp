<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->

<%

dim pagesize
dim notitemlist, itemlist
dim notitemlistinclude, itemlistinclude
dim research
dim yyyy1,mm1,dd1,yyyymmdd,nowdate
dim onlyOne,dcnt
dim danpumcheck
dim upbeaInclude
dim dcnt2
dim imsi, sagawa, ems, epostmilitary, bigitem
dim searchtypestring
dim deliveryarea
dim onejumuntype
dim onejumuncount, onejumuncompare
dim tenbeaonly



'==============================================================================
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

pagesize = request("pagesize")

deliveryarea = request("deliveryarea")

bigitem = request("bigitem")

upbeaInclude = request("upbeaInclude")

notitemlistinclude = request("notitemlistinclude")
itemlistinclude = request("itemlistinclude")

notitemlist = request("notitemlist")
itemlist = request("itemlist")

research = request("research")

onejumuntype = request("onejumuntype")
onejumuncount = request("onejumuncount")
onejumuncompare = request("onejumuncompare")

tenbeaonly = request("tenbeaonly")


'dcnt = trim(request("dcnt"))
'dcnt2 = trim(request("dcnt2"))
'onlyOne = request("onlyOne")
'danpumcheck = request("danpumcheck")
'imsi  = request("imsi")
'sagawa= request("sagawa")
'ems   = request("ems")
'epostmilitary   = request("epostmilitary")



'==============================================================================
if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2))-2,Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

if onejumuncount="" then
	onejumuncount = "1"
end if

if onejumuncompare="" then
	onejumuncompare = "less"
end if

if deliveryarea<>"" then
	if (deliveryarea = "ZZ") then
		ems   = ""
		epostmilitary   = "on"
	elseif (deliveryarea = "EMS") then
		ems   = "on"
		epostmilitary   = ""
	else
		deliveryarea = "KR"
		ems   = ""
		epostmilitary   = ""
	end if
end if

if (notitemlist = "") then
	notitemlistinclude = ""
end if

if (itemlist = "") then
	itemlistinclude = ""
end if

''�ӽ�..
if (research="") then
    notitemlist = "311341"
    notitemlistinclude="on"
end if

if research="" then
	'notitemlist = "45718"
	''if notitemlist="" then notitemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	''if itemlist="" then itemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	'if notitemlistinclude="" then notitemlistinclude="on"
end if

if (pagesize="") then
	pagesize = request.cookies("baljupagesize")
end if

if (pagesize="") then pagesize=200

response.cookies("baljupagesize") = pagesize






dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CTenBalju

''�� ����¡�� 2�� �˻�
ojumun.FPageSize = pagesize * 3

if notitemlistinclude="on" then
	ojumun.FRectNotitemlist = notitemlist
else
	ojumun.FRectNotitemlist = ""
end if

if itemlistinclude="on" then
	ojumun.FRectItemlist = itemlist
end if

ojumun.FCurrPage = page

ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1

''��ü��� ���� �ֹ���.
ojumun.FRectUpbeaInclude = upbeaInclude

''�簡�� ��۱ǿ�
ojumun.FRectOnlySagawaDeliverArea = sagawa

if deliveryarea<>"" then
	ojumun.FRectDeliveryArea = deliveryarea
end if

if bigitem<>"" then
	ojumun.FRectOnlyManyItem = "Y"
end if

if onejumuntype<>"" then
	ojumun.FRectOnlyOneJumun = "Y"

	ojumun.FRectOnlyOneJumunType = onejumuntype
	ojumun.FRectOnlyOneJumunCompare = onejumuncompare
	ojumun.FRectOnlyOneJumunCount = onejumuncount
end if

if tenbeaonly<>"" then
	ojumun.FRectTenbeaOnly = "Y"
end if


ojumun.GetBaljuItemListNew


dim ix,iy
dim tenbaljucount
tenbaljucount =0

%>
<script language='javascript'>
var tenBaljuCnt = 0;
function CheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;
    var isDasBalju = false;
    var isEmsBalju = <%= chkIIF(ems="on","true","false")%>;
    var isMilitaryBalju = <%= chkIIF(epostmilitary="on","true","false")%>;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� �ֹ��� �����ϴ�.');
		return;
	}

    if (document.all.groupform.songjangdiv.value.length<1){
		alert('��� �ù�縦 ���� �ϼ���.');
		document.all.groupform.songjangdiv.focus();
		return;
	}

	if (document.all.groupform.workgroup.value.length<1){
		alert('�۾� �׷��� ���� �ϼ���.');
		document.all.groupform.workgroup.focus();
		return;
	}
    //C�۾��� DAS
    isDasBalju = (document.all.groupform.workgroup.value=="C");

    //DAS ���� üũ, �ٹ� 150�� ����.
    if ((isDasBalju)&&(tenBaljuCnt>150)){
        alert('DAS ���ִ� �ٹ����� ��� 150�� �̸��� �����մϴ�. ');
		document.all.groupform.workgroup.focus();
		return;
    }

	// ========================================================================
    if (isEmsBalju){
        if (document.all.groupform.workgroup.value!="E"){
            alert('EMS���ִ� E �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="E"){
            alert('�˻������� �ؿܹ���̾�� EMS���ְ� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="90")&&(document.all.groupform.workgroup.value!="E"))||((document.all.groupform.songjangdiv.value!="90")&&(document.all.groupform.workgroup.value=="E"))){
        alert('EMS���ִ� E �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isMilitaryBalju){
        if (document.all.groupform.workgroup.value!="G"){
            alert('���δ� ���ִ� G �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="G"){
            alert('�˻������� ���δ����̾�� ���δ���ְ� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="8")&&(document.all.groupform.workgroup.value!="G"))||((document.all.groupform.songjangdiv.value!="8")&&(document.all.groupform.workgroup.value=="G"))){
        alert('���δ� ���ִ� G �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isDasBalju){
        if (!confirm('DAS ���� �Դϴ�. ��� �Ͻðڽ��ϱ�?')){
            return;
        }
    }
	var ret = confirm('���� �ֹ��� �� ���ּ��� �����Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.sitename.value = upfrm.sitename.value + "|" + frm.sitename.value;
				}
			}
		}
		upfrm.songjangdiv.value = document.all.groupform.songjangdiv.value;
		upfrm.workgroup.value = document.all.groupform.workgroup.value;
		upfrm.ems.value = "<%= ems %>";
		upfrm.epostmilitary.value = "<%= epostmilitary %>";

		if (isDasBalju) {
		    upfrm.baljutype.value = "D";
		}else{
		    upfrm.baljutype.value = "";
		}

		upfrm.submit();
	}
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function EnableDiable(icomp){
	//return;
	var frm = document.frm;
	var ischecked = icomp.checked;
	if (ischecked){
		if (icomp.name=="notitemlistinclude"){
			frm.itemlistinclude.checked = !(ischecked);
		}else if (icomp.name=="itemlistinclude"){
			frm.notitemlistinclude.checked = !(ischecked);
		}

	}

	if (icomp.name=="onlyOne"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.danpumcheck.checked = false;
	}


	if (icomp.name=="danpumcheck"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.onlyOne.checked = false;
	}
}

function poponeitem(){
	var popwin = window.open("/admin/ordermaster/poponeitem.asp","poponeitem","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chkUpbea(){
    var frm;
    var checkedExists = false;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.tenbeaexists.value!="Y"){
			    frm.cksel.checked = true;
			    AnCheckClick(frm.cksel);
			    checkedExists = true;
			}
		}
	}

	if (checkedExists){
	    document.groupform.songjangdiv.value="24";
	    document.groupform.workgroup.value="Z";
	    CheckNBalju();
	}
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr height="35">
        			<td width="250"><b>�Ⱓ : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ ����</td>
        			<td width="220">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>��� �Ǽ�</b> :
						<select name="pagesize" >
						<option value="10" <% if pagesize="10" then response.write "selected" %> >10</option>
						<option value="20" <% if pagesize="20" then response.write "selected" %> >20</option>
						<option value="50" <% if pagesize="50" then response.write "selected" %> >50</option>
						<option value="100" <% if pagesize="100" then response.write "selected" %> >100</option>
						<option value="150" <% if pagesize="150" then response.write "selected" %> >150</option>
						<option value="200" <% if pagesize="200" then response.write "selected" %> >200</option>
						<option value="250" <% if pagesize="250" then response.write "selected" %> >250</option>
						<option value="300" <% if pagesize="300" then response.write "selected" %> >300</option>
						<option value="400" <% if pagesize="400" then response.write "selected" %> >400</option>
						<option value="500" <% if pagesize="500" then response.write "selected" %> >500</option>
						<option value="600" <% if pagesize="600" then response.write "selected" %> >600</option>
						<option value="800" <% if pagesize="800" then response.write "selected" %> >800</option>
						<option value="1000" <% if pagesize="1000" then response.write "selected" %> >1000</option>
						<option value="2000" <% if pagesize="2000" then response.write "selected" %> >2000</option>
						</select>
        			</td>
        			<td width="200">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>�������</b> :
						<select name="deliveryarea" >
						<option value="" 	<% if deliveryarea="" then response.write "selected" %> >��ü</option>
						<option value="KR" 	<% if deliveryarea="KR" then response.write "selected" %> >�������</option>
						<option value="EMS" <% if deliveryarea="EMS" then response.write "selected" %> >�ؿܹ��</option>
						<option value="ZZ" 	<% if deliveryarea="ZZ" then response.write "selected" %> >���δ���</option>
						</select>
        			</td>
        			<td width="150">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="bigitem" <% if bigitem="on" then response.write "checked" %> > <b>�ټ���ǰ�ֹ�</b>
        			</td>
        			<td>&nbsp;</td>
        		</tr>
        		<tr height="35">
        			<td>
						<font color="#AAAAAA">
						<input type="checkbox" name="upbeaInclude" <% if upbeaInclude="on" then response.write "checked" %> > <b>�������� �ֹ��Ǹ�</b>
						</font>
        			</td>
        			<td colspan=4>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="tenbeaonly" <% if tenbeaonly="on" then response.write "checked" %> > <b>�ٹ��ֹ��Ǹ�</b>
        			</td>
        		</tr>
        		<tr height="35">
        			<td>
						<input type="checkbox" name="notitemlistinclude" <% if notitemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
						<b>Ư����ǰ ������ �ֹ�����</b> : <input type=text name=notitemlist value="<%= notitemlist %>" size=8 >
        			</td>
        			<td colspan=4>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="itemlistinclude" <% if itemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
						<b>Ư����ǰ ������ �ֹ��Ǹ�</b> : <input type=text name=itemlist value="<%= itemlist %>" size=8 >
        			</td>
        		</tr>
        		<tr height="35">
        			<td>
        				<b>��ǰ�ֹ�</b> :
						<select name="onejumuntype" >
						<option value="" 	<% if onejumuntype="" then response.write "selected" %> >========</option>
						<option value="all" <% if onejumuntype="all" then response.write "selected" %> >��� ��ǰ�ֹ�</option>
						<option value="reg" <% if onejumuntype="reg" then response.write "selected" %> >������ ��ǰ�ֹ�</option>
						</select>

						<input type="text" name="onejumuncount" value="<%= onejumuncount %>" size=3>
						<select name="onejumuncompare" >
						<option value="less" 	<% if onejumuncompare="less" then response.write "selected" %> >�� ����</option>
						<option value="more" 	<% if onejumuncompare="more" then response.write "selected" %> >�� �̻�</option>
						<option value="equal" 	<% if onejumuncompare="equal" then response.write "selected" %> >��</option>
						</select>

						<!--<input type="checkbox" name="onlyOne" <% if onlyOne="on" then response.write "checked" %> onclick="EnableDiable(this);">-->
        			</td>
        			<td colspan=4>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<!--<input type="checkbox" name="danpumcheck" <% if danpumcheck="on" then response.write "checked" %> onclick="EnableDiable(this);">-->
						<input type="button" value="��ǰ����ǰ����" onclick="javascript:poponeitem();">
						<!--<input type="text" name="dcnt2" value="<%= dcnt2 %>" size=1> �� (11 �Է½� 11�� �̻�, 0�� �Է½� 0�� �̻�)-->
        			</td>
        		</tr>

        	</table>




			<!--
			<input type="checkbox" name="ems" <% if ems="on" then response.write "checked" %> > <b>�ؿܹ��</b>

			<input type="checkbox" name="epostmilitary" <% if epostmilitary="on" then response.write "checked" %> > <b>���δ�</b>
			-->


			<!--
			<input type="checkbox" name="imsi" <% if imsi="on" then response.write "checked" %> > <b>�ӽ�(���ѵ��� ���� ����)</b>
			<font color="#AAAAAA">
			<input type="checkbox" name="sagawa" <% if sagawa="on" then response.write "checked" %> onClick="alert('�Ϲݹ��ָ� ���� (��ǰ���,Ư����ǰ �˻� ����ȵ�)');"> �ӽ�(�簡�ͱǿ�)
			</font>
			-->

        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        �� �̹��� �Ǽ� : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %></b></font>&nbsp;
			�� �ݾ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
			��հ��ܰ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font>
        </td>
        <td>&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="40" valign="center">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="200" align="left">
	        <div id="currsearchno">�� �˻��ֹ��Ǽ� : </div>
	        <div id="currtensearchno">�ٹ����ٹ�� �ֹ��Ǽ� : </div>
	        <!-- input type="checkbox" name="ck_upbea" onClick="chkUpbea();"> ������� -->
        </td>
        <td align="right">
		<form name="groupform">
		    <!--
		    <select name="baljutype">
		        <option value="">�Ϲ�
		        <option value="D">DAS
		        <option value="S">��ǰ���
		    </select>
		    -->
		    <select name="songjangdiv">
		        <option value="">�ù�缱��
<!--		   	<option value="2" >�����ù�	-->
                <% if (now()>"2010-04-01") then %>
                <option value="4" >CJ�ù�
                <% else %>
                    <option value="4" >CJ�ù�
			   	    <option value="24" >�簡��
			   	<% end if %>
			   	<option value="90" <%= CHKIIF(ems="on","selected","") %> >EMS
			   	<option value="8" <%= CHKIIF(epostmilitary="on","selected","") %> >��ü��(���δ�)
		    </select>
			<select name="workgroup">
			   	<option value="">�۾��׷�
			   	<option value="A" >A
			   	<option value="B" >B
			   	<option value="C" >C(DAS)
			   	<option value="D" >D
			   	<option value="F" >F
			   	<option value="E" <%= CHKIIF(ems="on","selected","") %> >E(EMS)
			   	<option value="G" <%= CHKIIF(epostmilitary="on","selected","") %> >G(���δ�)
			   	<option value="Z" >Z(����)
		   	</select>
			<input type="button" value="���û��׹��ּ��ۼ�" onclick="CheckNBalju()">
		</form>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20" align="center"><input type="checkbox" name="cksel" onClick="AnSelectAllFrame(true)"></td>
		<td width="80">�ֹ���ȣ</td>
		<td width="70">Site</td>
		<td width="70">����</td>
		<td width="80">UserID</td>
		<td width="80">������</td>
		<td width="80">������</td>
		<td width="60">�����ݾ�</td>
		<td width="60">�����Ѿ�</td>
		<td width="80">�������</td>
		<td width="80">�ŷ�����</td>
		<td width="110">�ֹ���</td>
		<td>
		    <% if upbeaInclude<>"" then %>
		    ��������
		    <% else %>
		    �ٹ�����
		    <% end if %>
		    </td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<% if tenbaljucount< CLng(pagesize) then %>
		<form name="frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>" method="post" >
		<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(ix).FOrderSerial %>">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="sitename" value="<%= ojumun.FItemList(ix).FSiteName %>">
		<input type="hidden" name="dlvcontrycode" value="<%= ojumun.FItemList(ix).FDlvcountryCode %>">

		<tr align="center" bgcolor="#FFFFFF">
		    <% if ((ems<>"") or (epostmilitary<>"")) then %>
		    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		    <% else %>
			<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <%= CHKIIF(ojumun.FItemList(ix).FDlvcountryCode<>"" and ojumun.FItemList(ix).FDlvcountryCode<>"KR","disabled","") %> ></td>
			<% end if %>
			<td><a href="javascript:ViewOrderDetail(frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
			<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
			<td><%= ojumun.FItemList(ix).FDlvcountryCode %></td>
			<td><%= ojumun.FItemList(ix).FUserID %></td>
			<td><%= ojumun.FItemList(ix).FBuyName %></td>
			<td><%= ojumun.FItemList(ix).FReqName %></td>
			<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
			<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
			<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
			<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
			<td><%= Left(ojumun.FItemList(ix).FRegDate,16) %></td>
			<td>

			<% if ojumun.FItemList(ix).Ftenbeaexists then %>
			<input type="hidden" name="tenbeaexists" value="Y">
			<% tenbaljucount = tenbaljucount + 1 %>
			�� <%= tenbaljucount %>
			<% else %>
            <input type="hidden" name="tenbeaexists" value="N">
			<% end if %>
			</td>
		</tr>
		</form>
	<% else %>
		<% exit for %>
	<% end if %>
	<% next %>
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<form name="frmArrupdate" method="post" action="dobaljumaker.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="sitename" value="">
<input type="hidden" name="songjangdiv" value="">
<input type="hidden" name="workgroup" value="">
<input type="hidden" name="baljutype" value="">
<input type="hidden" name="ems" value="">
<input type="hidden" name="epostmilitary" value="">
</form>
<%
set ojumun = Nothing
%>

<script language='javascript'>
document.all.currsearchno.innerHTML = "�˻����� : <Font color='#3333FF'><%= ix %></font>";
document.all.currtensearchno.innerHTML = "�ٹ����ٹ�� �˻����� : <Font color='#3333FF'><%= tenbaljucount %></font>";
tenBaljuCnt = 1*<%= tenbaljucount %>;
<% if onlyOne<>"" then %>
EnableDiable(frm.onlyOne);
<% end if %>

<% if danpumcheck<>"" then %>
EnableDiable(frm.danpumcheck);
<% end if %>

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->