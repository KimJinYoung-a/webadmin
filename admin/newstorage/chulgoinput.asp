<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �����Ʈ
' History : �̻� ����
'			2017.11.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim storeid, divcode, scheduledt, vatcode, chargeid, chargename, comment, storemarginrate
dim ArrShopInfo, currencyunit, currencyChar, loginsite, shopdiv, sqlStr, company_no, ischulgonotdisp
	storeid = request("storeid")
	divcode = request("divcode")
	scheduledt = request("scheduledt")
	vatcode = "008"
	chargeid = session("ssBctid")
	chargename = session("ssBctCname")
	comment = html2db(request("comment"))
	storemarginrate = request("storemarginrate")

ischulgonotdisp=false

if ((storeid <> "") and ((storemarginrate = "") or (storemarginrate = "0"))) then
	sqlStr = "select IsNull(a.marginrate, 0) as marginrate "
	sqlStr = sqlStr + " from [db_storage].[dbo].vw_acount_user_delivery a "
	sqlStr = sqlStr + " where a.userid = '" + storeid + "' "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		storemarginrate = rsget("marginrate")
	else
		storemarginrate = "0"
	end if
	rsget.close
elseif (storemarginrate = "") then
	storemarginrate = "0"
end if

if storeid<>"" then
	ArrShopInfo = getoffshopuser(storeid)

	IF isArray(ArrShopInfo) then
		currencyunit = ArrShopInfo(1,0)
		currencyChar = ArrShopInfo(3,0)
		loginsite = ArrShopInfo(2,0)
		shopdiv = ArrShopInfo(12,0)
    END IF

	sqlStr = "select id, company_no" & vbcrlf
	sqlStr = sqlStr & " from db_partner.dbo.tbl_partner" & vbcrlf
	sqlStr = sqlStr & " where id = '"& storeid &"'" & vbcrlf

    'response.write sqlStr & "<br>"
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		company_no = rsget("company_no")
	end if
	rsget.close
end if

' ���ų� �ؿ��ϰ�� �ٹ����� ����ڰ� �ƴҰ�� �̸Ŵ����� ������.
if Not(C_ADMIN_AUTH) and (replace(company_no,"-","")<>"2118700620" and (shopdiv="5" or shopdiv="7")) then
    ischulgonotdisp = true
end if

dim itemgubunarr, itemidarr, itemoptionarr
dim itemnamearr, itemoptionnamearr
dim sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr, mwdivarr

dim itemgubun, itemid, itemoption
dim itemname, itemoptionname
dim sellcash, suplycash, buycash, itemno, designer, mwdiv

itemgubunarr = request("itemgubunarr")
itemidarr	= request("itemidarr")
itemoptionarr = request("itemoptionarr")
itemnamearr		= request("itemnamearr")
itemoptionnamearr = request("itemoptionnamearr")
sellcasharr = request("sellcasharr")
suplycasharr = request("suplycasharr")
buycasharr = request("buycasharr")
itemnoarr = request("itemnoarr")
designerarr = request("designerarr")
mwdivarr = request("mwdivarr")

%>
<script>
function Items2Array()
{
	var frm;

	frmMaster.itemgubunarr.value = "";
	frmMaster.itemidarr.value = "";
	frmMaster.itemoptionarr.value = "";
	frmMaster.itemnamearr.value = "";
	frmMaster.itemoptionnamearr.value = "";
	frmMaster.sellcasharr.value = "";
	frmMaster.suplycasharr.value = "";
	frmMaster.buycasharr.value = "";
	frmMaster.itemnoarr.value = "";
	frmMaster.designerarr.value = "";
	frmMaster.mwdivarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (!IsInteger(frm.itemno.value)){
				alert('������ ������ �����մϴ�.');
				frm.itemno.focus();
				return;
			}

			if (!IsInteger(frm.suplycash.value)){
				alert('����� ������ �����մϴ�.');
				frm.suplycash.focus();
				return;
			}

			frmMaster.itemgubunarr.value = frmMaster.itemgubunarr.value + frm.itemgubun.value + "|";
			frmMaster.itemidarr.value = frmMaster.itemidarr.value + frm.itemid.value + "|";
			frmMaster.itemoptionarr.value = frmMaster.itemoptionarr.value + frm.itemoption.value + "|";
			frmMaster.itemnamearr.value = frmMaster.itemnamearr.value + frm.itemname.value + "|";
			frmMaster.itemoptionnamearr.value = frmMaster.itemoptionnamearr.value + frm.itemoptionname.value + "|";
			frmMaster.sellcasharr.value = frmMaster.sellcasharr.value + frm.sellcash.value + "|";
			frmMaster.suplycasharr.value = frmMaster.suplycasharr.value + frm.suplycash.value + "|";
			frmMaster.buycasharr.value = frmMaster.buycasharr.value + frm.buycash.value + "|";
			frmMaster.itemnoarr.value = frmMaster.itemnoarr.value + (frm.itemno.value * -1) + "|";
			frmMaster.designerarr.value = frmMaster.designerarr.value + frm.desingerid.value + "|";
			frmMaster.mwdivarr.value = frmMaster.mwdivarr.value + frm.mwdiv.value + "|";
		}
	}

}

function removeDuplicate() {
	var itemgubunarr, itemidarr, itemoptionarr, itemnamearr, itemoptionnamearr, sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr, mwdivarr;
	var i, j;

	itemgubunarr = frmMaster.itemgubunarr.value.split("|");
	itemidarr = frmMaster.itemidarr.value.split("|");
	itemoptionarr = frmMaster.itemoptionarr.value.split("|");
	itemnamearr = frmMaster.itemnamearr.value.split("|");
	itemoptionnamearr = frmMaster.itemoptionnamearr.value.split("|");
	sellcasharr = frmMaster.sellcasharr.value.split("|");
	suplycasharr = frmMaster.suplycasharr.value.split("|");
	buycasharr = frmMaster.buycasharr.value.split("|");
	itemnoarr = frmMaster.itemnoarr.value.split("|");
	designerarr = frmMaster.designerarr.value.split("|");
	mwdivarr = frmMaster.mwdivarr.value.split("|");

	frmMaster.itemgubunarr.value = "";
	frmMaster.itemidarr.value = "";
	frmMaster.itemoptionarr.value = "";
	frmMaster.itemnamearr.value = "";
	frmMaster.itemoptionnamearr.value = "";
	frmMaster.sellcasharr.value = "";
	frmMaster.suplycasharr.value = "";
	frmMaster.buycasharr.value = "";
	frmMaster.itemnoarr.value = "";
	frmMaster.designerarr.value = "";
	frmMaster.mwdivarr.value = "";

	for (i = 0; i < itemgubunarr.length; i++) {
		if ((itemgubunarr[i] != "XX") && (itemgubunarr[i] != "")) {
			for (j = i + 1; j < itemgubunarr.length; j++) {
				if ((itemgubunarr[i] == itemgubunarr[j]) && (itemidarr[i] == itemidarr[j]) && (itemoptionarr[i] == itemoptionarr[j])) {
					itemgubunarr[j] = "XX";
					itemnoarr[i] = itemnoarr[i]*1 + itemnoarr[j]*1;
				}
			}

			frmMaster.itemgubunarr.value = frmMaster.itemgubunarr.value + itemgubunarr[i] + "|";
			frmMaster.itemidarr.value = frmMaster.itemidarr.value + itemidarr[i] + "|";
			frmMaster.itemoptionarr.value = frmMaster.itemoptionarr.value + itemoptionarr[i] + "|";
			frmMaster.itemnamearr.value = frmMaster.itemnamearr.value + itemnamearr[i] + "|";
			frmMaster.itemoptionnamearr.value = frmMaster.itemoptionnamearr.value + itemoptionnamearr[i] + "|";
			frmMaster.sellcasharr.value = frmMaster.sellcasharr.value + sellcasharr[i] + "|";
			frmMaster.suplycasharr.value = frmMaster.suplycasharr.value + suplycasharr[i] + "|";
			frmMaster.buycasharr.value = frmMaster.buycasharr.value + buycasharr[i] + "|";
			frmMaster.itemnoarr.value = frmMaster.itemnoarr.value + itemnoarr[i] + "|";
			frmMaster.designerarr.value = frmMaster.designerarr.value + designerarr[i] + "|";
			frmMaster.mwdivarr.value = frmMaster.mwdivarr.value + mwdivarr[i] + "|";
		}
	}
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,imwdiv){
	if (iidx!='0'){
		alert('�ֹ����� ��ġ���� �ʽ��ϴ�. �ٽýõ��� �ּ���.');
		return;
	}

    //��� �⺻ 0��ó��
    var arrsuplycash = isuplycash.split("|");
    isuplycash = "";
    for (var i=0;i<arrsuplycash.length;i++){
        if(i==0){
            isuplycash =  parseInt(arrsuplycash[i])*0;
        }else{
        isuplycash = isuplycash + "|" + parseInt(arrsuplycash[i])*0;
        }
    }

	Items2Array();

	frmMaster.itemgubunarr.value = frmMaster.itemgubunarr.value + igubun;
	frmMaster.itemidarr.value = frmMaster.itemidarr.value + iitemid;
	frmMaster.itemoptionarr.value = frmMaster.itemoptionarr.value + iitemoption;
	frmMaster.sellcasharr.value = frmMaster.sellcasharr.value + isellcash;
	frmMaster.suplycasharr.value = frmMaster.suplycasharr.value + isuplycash;
	//frmMaster.suplycasharr.value = frmMaster.suplycasharr.value + isellcash;

	frmMaster.buycasharr.value = frmMaster.buycasharr.value + ibuycash;
	frmMaster.itemnoarr.value = frmMaster.itemnoarr.value + iitemno;
	frmMaster.itemnamearr.value = frmMaster.itemnamearr.value + iitemname;
	frmMaster.itemoptionnamearr.value = frmMaster.itemoptionnamearr.value + iitemoptionname;
	frmMaster.designerarr.value = frmMaster.designerarr.value + iitemdesigner;
	frmMaster.mwdivarr.value = frmMaster.mwdivarr.value + imwdiv;

	removeDuplicate();

	frmMaster.submit();
}

function AddItems(frm){
	var popwin;
	var suplyer, shopid;
	var frm = document.frmMaster;
	var priceGbn;

	if (frm.storeid.value === "") {
		alert("���� ���ó�� �Է��ϼ���.");
		return;
	}

	if (frm.storeid.value === "itemgift") {
		// ���������� �δ� ���رݾ��� ���ΰ�
		priceGbn = "&priceGbn=saleprice"
	} else {
		priceGbn = "&priceGbn=orgprice"
	}
	popwin = window.open('/admin/newstorage/popjumunitemNew.asp?suplyer=&changesuplyer=Y&shopid=10x10&idx=0' + priceGbn,'chulgoinputadd','width=1280,height=960,scrollbars=yes,resizable=no');
	popwin.focus();
}

function ApplyMargin() {
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			frm.suplycash.value = 1 * frm.sellcash.value * (100 - frmMaster.storemarginrate.value) / 100;
		}
	}
}

function SubmitForm() {
	var frm = document.frmMaster;

    if (frm.storeid.value == "") {
        alert("���ó�� �����ϼ���.");
        return;
    }

	if ( (frm.storeid.value == "promotion") ) {		//  || (frm.storeid.value == "etcsales")
		alert("���ó promotion �� ������ �� �����ϴ�.");
		//alert("���ó promotion, etcsales �� ������ �� �����ϴ�.");
        return;
	}

	if ( (frm.divcode.value == "999") ) {
		alert("����� ��Ÿ(�������)��  ������ �� �����ϴ�.");
        return;
	}

    if (frm.divcode.value == "") {
        alert("������� �����ϼ���.");
        return;
    }
    if (frm.vatcode.value == "") {
        alert("�ΰ��������� �����ϼ���.");
        return;
    }
    if (frm.scheduledt.value == "") {
        alert("����û���� �Է��ϼ���.");
        return;
    }

    if (confirm("�����Ͻðڽ��ϱ�?") != true) {
        return;
	}

    Items2Array();

    frm.mode.value = "write";
    frm.action = "chulgoedit_process.asp";
    frm.submit();

}

function tempSave(){
	var frm = document.frmMaster;

	if (frm.storeid.value == "") {
        alert("���ó�� �����ϼ���.");
        return;
    }

	if ( (frm.storeid.value == "promotion") ) {		//  || (frm.storeid.value == "etcsales")
		alert("���ó promotion �� ������ �� �����ϴ�.");
		//alert("���ó promotion, etcsales �� ������ �� �����ϴ�.");
        return;
	}

	if ( (frm.divcode.value == "999") ) {
		alert("����� ��Ÿ(�������)��  ������ �� �����ϴ�.");
        return;
	}

    Items2Array();

	frm.mode.value = "temp";
    frm.action = "chulgoedit_process.asp";
    frm.submit();
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
        	<font color="red"><strong>����Է�</strong></font>
		</td>
	</tr>
	<!-- ��ܹ� �� -->

	<form name="frmMaster" method="post" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="chargeid" value="<%= chargeid %>">
	<input type="hidden" name="chargename" value="<%= chargename %>">
	<input type="hidden" name="vatcode" value="<%= vatcode %>">

	<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
	<input type="hidden" name="itemidarr" value="<%= itemidarr %>">
	<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
	<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
	<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
	<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
	<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
	<input type="hidden" name="buycasharr" value="<%= buycasharr %>">
	<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
	<input type="hidden" name="designerarr" value="<%= designerarr %>">
	<input type="hidden" name="mwdivarr" value="<%= mwdivarr %>">
    <tr align="center" bgcolor="#FFFFFF">
		<td width=100 bgcolor="<%= adminColor("tabletop") %>">���ó</td>
		<td width=400 align="left">	<% drawSelectBoxOffShopNotUsingAll "storeid",storeid %> <!--% drawSelectBoxChulgo "storeid", storeid %--></td>
		<td width=100 bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td align="left">
			<% Call drawSelectBoxIpChulDivcode("etcchulgo", "divcode", divcode) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">����û��</td>
		<td align="left"><input type="text" class="text" name="scheduledt" value="<%= scheduledt %>" size="10" maxlength=10 readonly><a href="javascript:calendarOpen(frmMaster.scheduledt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a></td>
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td align="left"><%= chargename %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
		<td colspan="3" align="left"><textarea class="textarea" name="comment" cols=80 rows=6><%= comment %></textarea></td>
	</tr>
</table>
<%

itemgubunarr = split(itemgubunarr,"|")
itemidarr	= split(itemidarr,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")
mwdivarr = split(mwdivarr,"|")

dim cnt, i

cnt = ubound(itemidarr)
if cnt < 0 then cnt = 0
dim selltotal, suplytotal, buytotal
selltotal = 0
suplytotal = 0
buytotal = 0

%>

<br>
<font color="blue">+ ����� �⺻������ 0������ ��ϵ˴ϴ�. ������ ���Ͻø� ��ǰ�߰� �� [�������ϰ�����]��ư�� �̿��ؼ� ����� �������ּ���</font>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="9">
			<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
			        	<font color="red"><strong>�󼼳���</strong></font>
			        	&nbsp;&nbsp;
			        	<font color="#EE4444">����</font>&nbsp;��Ź&nbsp;<font color="#4444EE">��ü���</font>
			        	&nbsp;&nbsp;
			        	 ���������:
			        	<input type="text" class="text" style="text-align:right;" name="storemarginrate" value="<%= storemarginrate %>" size="2"> %
			        	<input type="button" class="button" value="�������ϰ�����" onclick="ApplyMargin()">

	        		</td>
	        		<td align="right">
	        			�ѰǼ�:  <%= cnt %>
			        	&nbsp;
			        	<input type="button" class="button" value=" ��ǰ�߰� " onClick="AddItems(frmMaster)" <% if ischulgonotdisp then %> disabled<% end if %>>
	        		</td>
	        	</tr>
	        </table>
		</td>
	</tr>
	</form>
	<!-- ��ܹ� �� -->

    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="90">���ڵ�</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td width="80">�ǸŰ�</td>
		<td width="80">���</td>
		<td width="80">���԰�</td>
		<td width="60">����</td>
		<td width="50">������</td>
		<td width="50">���Ը���</td>
		<!--
		<td width="60">����</td>
		-->
	</tr>
	<% for i=0 to cnt-1 %>
	<%
	selltotal  = selltotal + sellcasharr(i) * itemnoarr(i)
	suplytotal = suplytotal + suplycasharr(i) * itemnoarr(i)
	buytotal = buytotal + buycasharr(i) * itemnoarr(i)
	%>
	<form name="frmBuyPrc_<%= i %>" method="post" action="">
	<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
	<input type="hidden" name="itemid" value="<%= itemidarr(i) %>">
	<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
	<input type="hidden" name="itemname" value="<%= itemnamearr(i) %>">
	<input type="hidden" name="itemoptionname" value="<%= itemoptionnamearr(i) %>">
	<input type="hidden" name="desingerid" value="<%= designerarr(i) %>">
	<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
	<input type="hidden" name="mwdiv" value="<%= mwdivarr(i) %>">

	<tr bgcolor="#FFFFFF">
		<td align=center >
		<% if mwdivarr(i)="M" then %>
		<font color="#EE4444"><%= itemgubunarr(i) %>-<%= CHKIIF(itemidarr(i)>=1000000,format00(8,itemidarr(i)),format00(6,itemidarr(i))) %>-<%= itemoptionarr(i) %></font>
		<% elseif mwdivarr(i)="U" then %>
		<font color="#4444EE"><%= itemgubunarr(i) %>-<%= CHKIIF(itemidarr(i)>=1000000,format00(8,itemidarr(i)),format00(6,itemidarr(i))) %>-<%= itemoptionarr(i) %></font>
		<% else %>
		<%= itemgubunarr(i) %>-<%= CHKIIF(itemidarr(i)>=1000000,format00(8,itemidarr(i)),format00(6,itemidarr(i))) %>-<%= itemoptionarr(i) %>
		<% end if %>
		</td>
		<td ><%= itemnamearr(i) %></td>
		<td ><%= itemoptionnamearr(i) %></td>
		<td align=right><%= FormatNumber(sellcasharr(i),0) %></td>
		<td align=right><input type="text" class="text" name="suplycash" value="<%= suplycasharr(i) %>" size=7 maxlength=7></td>
		<td align=right><input type="text" class="text" name="buycash" value="<%= buycasharr(i) %>" size=7 maxlength=7></td>

		<td align=right><input type="text" class="text" name="itemno" value="<%= itemnoarr(i)*-1 %>"  size="4" maxlength="4"></td>
		<td align=center>
		<% if sellcasharr(i)<>0 then %>
			<%= 100-CLng(suplycasharr(i)/sellcasharr(i)*100*100)/100 %>%
		<% end if %>
		</td>
		<td align=center>
		<% if sellcasharr(i)<>0 then %>
			<%= 100-CLng(buycasharr(i)/sellcasharr(i)*100*100)/100 %>%
		<% end if %>
		</td>

		<!--
		<td>
		      <select name="mwdiv">
		        <option value="M" <% if (mwdivarr(i) = "M") then response.write "selected" end if %>>����</option>
		        <option value="W" <% if (mwdivarr(i) = "W") then response.write "selected" end if %>>��Ź</option>
		      </select>

	        </td>
	        -->
	</tr>
	</form>
	<% next %>

	<% if (cnt>0) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">�Ѱ�</td>
		<td colspan="2" align="center">
		<td align=right><%= formatNumber(selltotal,0) %></td>
		<td align=right><%= formatNumber(suplytotal,0) %></td>
		<td align=right><%= formatNumber(buytotal,0) %></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<% end if %>

</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1">
	<tr height="25"  >
		<td colspan="15" align="center">
			<% if ischulgonotdisp then %>
				<font color="red">���ó�� �ؿܳ� ���ŷ� �����Ǿ� �ִ°��, [OFF]����_������>>�ֹ�����(����)���� ��� �ϼž� �մϴ�.</font><Br>
			<% end if %>
			<input type="button" class="button" value="�ӽ�����(�ۼ���)" onclick="tempSave()" <% if ischulgonotdisp then %> disabled<% end if %>>
			<% if (cnt>0) then %>
			<input type="button" class="button" value="����Ȯ��(����)" onclick="SubmitForm()" <% if ischulgonotdisp then %> disabled<% end if %>>
        	<% else %>
        	&nbsp;
        	<% end if %>
		</td>
	</tr>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
