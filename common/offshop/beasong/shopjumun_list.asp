<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->
<%
dim showshopselect, loginidshopormaker ,tmodlvType ,ojumun , i , orderno, UserHpAuto, certsendgubun, odlvTypedefault, totrealprice
	orderno = requestcheckvar(request("orderno"),16)
	UserHpAuto = requestcheckvar(request("UserHpAuto"),16)
	certsendgubun = requestcheckvar(request("certsendgubun"),32)

totrealprice=0
showshopselect = false
loginidshopormaker = ""
if certsendgubun="" then certsendgubun = "KAKAOTALK"
odlvTypedefault="1"

if C_ADMIN_USER or C_IS_OWN_SHOP then
	showshopselect = true
	loginidshopormaker = request("shopid")
elseif (C_IS_SHOP) then
	'����/������
	loginidshopormaker = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		loginidshopormaker = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'ǥ�þ��Ѵ�. ����.
		else
			showshopselect = true
			loginidshopormaker = request("shopid")
		end if
	end if
end if

set ojumun = new cupchebeasong_list
	ojumun.frectorderno = orderno
	ojumun.frectshopid = loginidshopormaker

if (orderno <> "") then
	ojumun.fshopjumun_list()
end if

'//üũ�ڽ� disabled ���� ó��
function checkblock(CurrState)
	checkblock = false

	'// ����Է��� �ִ� ��� true
	if Not IsNull(CurrState) then
		checkblock = true
	end if
end Function
%>

<script language="javascript">

	//�ֹ�����
	function jumundetail(masteridx){
		frm.masteridx.value=masteridx;
		frm.action='/common/offshop/beasong/shopbeasong_input.asp';
		frm.submit();
	}

	// ����� �ϰ� ����
	function chodlvTypedefault(){
		var odlvTypedefault = document.getElementById("odlvTypedefault");		//getElementsByName

		if (odlvTypedefault.value==''){
			alert('�ϰ����� �Ͻ� ���� ����ڰ� �����Ǿ� ���� �ʽ��ϴ�.');
			odlvTypedefault.focus();
			return;
		}

		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				frm.cksel.checked = true;
				AnCheckClick(frm.cksel);
				frm.odlvType.value=odlvTypedefault.value;
			}
		}
	}
	
	//���忡�� ��ǰ ��� ��û
	function jumuninput(upfrm,isAuto){
		if (upfrm.orderno.value==''){
			alert('�ֹ���ȣ�� �Է� �ϼ���');
			upfrm.orderno.focus();
			return;
		}
		var orderno = upfrm.orderno.value;

		if (isAuto=='auto'){
			if (upfrm.certsendgubun.value==''){
				alert('�ڵ�����ȣ�� �߼��� ���� ������ �����ϴ�.');
				upfrm.certsendgubun.focus();
				return;
			}
			var certsendgubun = upfrm.certsendgubun.value;

			if (upfrm.UserHpAuto.value==''){
				alert('���� ������ �޴�����ȣ�� �Է� �ϼ���');
				upfrm.UserHpAuto.focus();
				return;
			}
			var UserHpAuto = upfrm.UserHpAuto.value;
		}

		upfrm.itemgubunarr.value='';
		upfrm.itemidarr.value='';
		upfrm.itemoptionarr.value='';
		upfrm.masteridx.value='';
		upfrm.masteridxarr.value='';
		upfrm.odlvTypearr.value='';

		if (!CheckSelected()){
			alert('���� ��ǰ�� �����ϴ�.');
			return;
		}

		var frm; var odlvType='';
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					if (frm.odlvType.value==''){
						alert('��۱����� ���� �ϼ���');
						frm.odlvType.focus();
						return;
					}
					// comm_cd : B031 ����������� / B012 ��üƯ�� / B013 ���Ư��
					// �������
					if (frm.odlvType.value*1 == '1') {
						if (frm.comm_cd.value == 'B012') {
							alert("�ش� ��ǰ�� ��ü��� or �����۸� ���� �մϴ�.");
							frm.odlvType.focus();
							return;
						}
					}
					// ��ü���
					if (frm.odlvType.value*1 == '2') {
						if (frm.comm_cd.value == 'B031' || frm.comm_cd.value == 'B013') {
							alert("�ش� ��ǰ�� ������� or �����۸� ���� �մϴ�.");
							frm.odlvType.focus();
							return;
						}
					}
					
/*
					if (frm.defaultbeasongdiv.value*1 == 0) {
						if (frm.odlvType.value*1 != 0) {
							alert("�����Ҽ� ���� ������Դϴ�. �������� �����ϼ���.");
							frm.odlvType.focus();
							return;
						}
					}

					if (odlvType!='' && odlvType != frm.odlvType.value){
						alert('��۱����� ��ǰ���� �ٸ��� �����ϽǼ� �����ϴ�.');
						frm.odlvType.focus();
						return;
					}
*/

					upfrm.odlvTypearr.value = upfrm.odlvTypearr.value + frm.odlvType.value + "," ;
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "," ;
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "," ;
					upfrm.shopidarr.value = upfrm.shopidarr.value + frm.shopid.value + "," ;

					// ����� ������ ���� ��� ������ �����Ѵ�.
					odlvType = frm.odlvType.value;
				}
			}
		}

		if (isAuto=='auto'){
			if (confirm('����( ' + UserHpAuto + ' )�� �ּ��Է� ��ũ�� '+ certsendgubun + ' ���� �߼� �Ͻðڽ��ϱ�?')){
				upfrm.mode.value='userjumun';
				upfrm.action='/common/offshop/beasong/shopbeasong_process.asp';
				upfrm.submit();
			}
		}else{
			if (confirm('�ּҸ� ����� ���忡�� ���� �Է� �Ͻðڽ��ϱ�?')){
				upfrm.mode.value='shopjumun';
				upfrm.action='/common/offshop/beasong/shopjumun_address.asp';
				upfrm.submit();
			}
		}
	}

	//���ε�� ����Ʈ
	function getOnload(){
	    frm.orderno.select();
	    frm.orderno.focus();
	    
	    <% if (session("poslogin") = 1) then %>
	    // POS���� �Ѿ�� ���̸�
	        setTimeout(function(){
	            reqPosSign();
            },100);
	    <% end if %>
	}

	window.onload = getOnload;

	//������
	function gosubmit(){
		frm.submit();
	}

	function CheckThis(frm){
		frm.cksel.checked=true;
		AnCheckClick(frm.cksel);
	}

	// �ʱⰪ ��ü ����
	function AllCheck(){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if ( frm.cksel.disabled != true){
					frm.cksel.checked = true;
					AnCheckClick(frm.cksel);
				}
			}
		}
	}

    /* pos ��� */
    function reqPosSign(){
        window.location = "jsTenPosCall://tenPosReqSign?resultPosSign";
    }
    
    function resultPosSign(ival){
        document.frm.UserHpAuto.value=ival;
    }
</script>

<!-- �˻� ���� -->
<form name="frm" method="post" action="">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="shopidarr">
<input type="hidden" name="masteridxarr">
<input type="hidden" name="mode">
<input type="hidden" name="masteridx">
<input type="hidden" name="odlvTypearr">
<input type="hidden" name="detailidxarr">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���� :
		<% if (showshopselect = true) then %>
			<% 'drawSelectBoxOffShop "shopid",loginidshopormaker %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",loginidshopormaker, "21") %>
		<% else %>
			<%= loginidshopormaker %>
		<% end if %>
		&nbsp;&nbsp;
		* �ֹ���ȣ : <input type="text" name="orderno" value="<%= orderno %>" size="16" onKeyPress="if(window.event.keyCode==13) gosubmit('');">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</table>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		* �޴�����ȣ : <input type="text" name="UserHpAuto" value="<%= UserHpAuto %>" size=16 maxlength=16>
		<% drawcertsendgubun "certsendgubun", certsendgubun, "", "N" %>
		<input type="button" value="���û�ǰ �ּ��Է� ��ũ�߼�" class="button" onclick="jumuninput(frm,'auto')">
		&nbsp;&nbsp;&nbsp;
		<input type="button" value="POS �����е��û"  class="button" onclick="reqPosSign()">
		
	</td>
	<td align="right">
		<input type="button" value="���û�ǰ �ּ��Է�(���������Է�)" class="button" onclick="jumuninput(frm,'')">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ojumun.FTotalCount %></b><br>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>�ǸŸ���</td>
	<td>�ֹ���ȣ</td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��ID</td>
	<td>��ǰ��<br>�ɼǸ�</td>
	<td>�Ǹűݾ�</td>
	<td>�ǰ�����</td>
	<td>�Ǹż���</td>
	<td>�հ�</td>
	<td>�Ǹ���</td>
	<td>�⺻��۱���</td>
	<td>
		���������
		<!--<Br><% 'Drawbeasonggubun "odlvTypedefault", odlvTypedefault," id='odlvTypedefault'" %>
		<input type="button" value="�ϰ�����" class="button" onclick="chodlvTypedefault();">-->
	</td>
	<td>��ۻ���</td>
	<td>���</td>
</tr>
<% if ojumun.FTotalCount>0 then %>
<% for i=0 to ojumun.FTotalCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="orderno" value="<%= ojumun.FItemList(i).forderno %>">
<input type="hidden" name="itemgubun" value="<%= ojumun.FItemList(i).fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ojumun.FItemList(i).fitemid %>">
<input type="hidden" name="itemoption" value="<%= ojumun.FItemList(i).fitemoption %>">
<input type="hidden" name="shopid" value="<%= ojumun.FItemList(i).fshopid %>">
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(i).fmasteridx %>">
<input type="hidden" name="detailidx" value="<%= ojumun.FItemList(i).fdetailidx %>">
<input type="hidden" name="comm_cd" value="<%= ojumun.FItemList(i).fcomm_cd %>">

<% if ojumun.FItemList(i).fcurrstate = "" or isnull(ojumun.FItemList(i).fcurrstate) then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="#FFFFaa">
<% end if %>
	<td>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if checkblock(ojumun.FItemList(i).FCurrState) then response.write " disabled" %>>
	</td>
	<td>
		<%= ojumun.FItemList(i).fshopname %>
	</td>
	<td>
		<%= ojumun.FItemList(i).forderno %>
	</td>
	<td>
		<%=ojumun.FItemList(i).fitemgubun%>-<%=CHKIIF(ojumun.FItemList(i).fitemid>=1000000,Format00(8,ojumun.FItemList(i).fitemid),Format00(6,ojumun.FItemList(i).fitemid))%>-<%=ojumun.FItemList(i).fitemoption%>
	</td>
	<td>
		<%=ojumun.FItemList(i).fmakerid%>
	</td>
	<td>
		<%= ojumun.FItemList(i).fitemname %><br><%= ojumun.FItemList(i).fitemoptionname %>
	</td>
	<td><%= FormatNumber(ojumun.FItemList(i).fsellprice,0) %></td>
	<td><%= FormatNumber(ojumun.FItemList(i).frealsellprice,0) %></td>
	<td>
		<%= ojumun.FItemList(i).fitemno %>
	</td>
	<td><%= FormatNumber(ojumun.FItemList(i).frealsellprice*ojumun.FItemList(i).fitemno,0) %></td>
	<td>
		<%= ojumun.FItemList(i).fIXyyyymmdd %>
	</td>
	<input type="hidden" name="defaultbeasongdiv" value="<%= ojumun.FItemList(i).Fdefaultbeasongdiv %>">
	<td>
		<% if (ojumun.FItemList(i).Fdefaultbeasongdiv <> 0) then %>
			<%= ojumun.FItemList(i).getDefaultBeasongDivName %>
		<% end if %>
	</td>
	<%
	tmodlvType = ojumun.FItemList(i).fodlvType

	if (tmodlvType = "") or (IsNull(tmodlvType)) then
		' �������(����������� , ���Ư��) , ��ü���(��üƯ��) , ������(ALL)
		if ojumun.FItemList(i).fcomm_cd="B031" or ojumun.FItemList(i).fcomm_cd="B013" then
			tmodlvType = "1"
		elseif ojumun.FItemList(i).fcomm_cd="B012" then
			tmodlvType = "2"
		else
			tmodlvType = "1"
		end if
	end if
	%>
	<td>
		<% Drawbeasonggubun "odlvType", tmodlvType," onchange='CheckThis(frmBuyPrc"& i &");'" %>
	</td>
	<td>
		<%= ojumun.FItemList(i).shopNormalUpcheDeliverState %>
		<%
		'//�����Ϸ� ���°� �ƴ϶��  ����û ��ȣ ������
		if ojumun.FItemList(i).FCurrState<>"" or not isnull(ojumun.FItemList(i).FCurrState) then
		%>
			<!--
			<br>(�ϷĹ�ȣ : <%= ojumun.FItemList(i).fmasteridx %>)
			-->
		<% end if %>
	</td>
	<td>
		<%
		'//�ֹ���� ���� �϶� �ּ� ���� ����
		if ojumun.FItemList(i).FCurrState="0" then
		%>
			<input type="button" onclick="jumundetail(<%= ojumun.FItemList(i).fmasteridx %>);" value="�ֹ�����" class="button">
		<%
		end if
		%>
	</td>
</tr>
</form>
<%
totrealprice = totrealprice + (ojumun.FItemList(i).frealsellprice*ojumun.FItemList(i).fitemno)
next
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=9>�հ�</td>
	<td><%= FormatNumber(totrealprice,0) %></td>
	<td colspan=9></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<script type="text/javascript">
	<% if ojumun.FTotalCount>0 then %>
		AllCheck();
	<% end if %>
</script>

<br>

* 200 �� ���� �˻� �˴ϴ�.<br>
* ��������� ���Է��Ͻ÷��� ���� �Էµ� ��������� �����ϼ���.
<%
set ojumun = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->