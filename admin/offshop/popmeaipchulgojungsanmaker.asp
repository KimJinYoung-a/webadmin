<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim page
dim shopid, startdate, enddate, ctype, onlymifinish, research
ctype = requestCheckVar(request("ctype"),10)
shopid = requestCheckVar(request("shopid"),32)
page = requestCheckVar(request("page"),10)
onlymifinish = requestCheckVar(request("onlymifinish"),2)
research = requestCheckVar(request("research"),2)

if ctype="" then ctype = "J"
if page="" then page = 1
if (research="") and (onlymifinish="") then onlymifinish="on"

dim nowdate, yyyy1,yyyy2,mm1,mm2,dd1,dd2

yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)
yyyy2 = requestCheckVar(request("yyyy2"),4)
mm2 = requestCheckVar(request("mm2"),2)
dd2 = requestCheckVar(request("dd2"),2)

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)

	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2) - 3
	dd1   = "01" ''Mid(nowdate,9,2)

	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

startdate = CStr(DateSerial(yyyy1 , mm1 , dd1))
enddate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

dim ofranchulgojungsan

set ofranchulgojungsan = new CFranjungsan
	ofranchulgojungsan.FPageSize=500
	ofranchulgojungsan.FCurrpage = page
	ofranchulgojungsan.FRectshopid = shopid
	ofranchulgojungsan.FRectStartDate = startdate
	ofranchulgojungsan.FRectEndDate = enddate
	ofranchulgojungsan.FRectonlymifinish = onlymifinish

	if shopid<>"" then
		if ctype="M" then
		''������� - ������
			ofranchulgojungsan.getChulgoJungsanTargetList
		elseif ctype="J" then
		''������� - �ֹ�����
			ofranchulgojungsan.getChulgoJungsanTargetListByJumun
		elseif ctype="W" then
		''��üƯ��
			ofranchulgojungsan.getWitakSellJungsanTargetList
		end if
	end if

dim i
dim ttlsell,ttlsuply,ttlbuy
ttlsell = 0
ttlsuply = 0
ttlbuy = 0
%>
<script type='text/javascript'>

function reCalcuSum(frm){
	var suplysum = 0;

	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			if (e.checked){
				suplysum = suplysum + eval("frm.val_" + e.value).value*1;
			}
		}
	}

	document.buffrm.totalsuply.value = suplysum;
}

function PopChulgoSheet(v){
	var popwin;
	popwin = window.open('/admin/newstorage/popchulgosheet.asp?idx=' + v ,'popchulgosheet','width=760,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=640,height=540,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopDetail(iidx,shopid){
	var popwin = window.open("/admin/offupchejungsan/off_jungsandetailsum.asp?gubuncd=B012&idx=" + iidx + '&shopid=' + shopid,"popjungsandetail","width=1000, height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function SaveArr(frm){
	var ischecked = false;
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			ischecked = (ischecked || e.checked);
			if (ischecked) break;
		}
	}

	if (!ischecked) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	var val_workidx = "-";
	var is_multiworkidx = false;
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			if (e.checked){
				if (val_workidx == "-") {
					val_workidx = eval("frm.val_workidx_" + e.value).value;
				}

				if (eval("frm.val_workidx_" + e.value).value != val_workidx) {
					is_multiworkidx = true;
					val_workidx = eval("frm.val_workidx_" + e.value).value;
				}
			}
		}
	}

	if (is_multiworkidx == true) {
		if (confirm("�̹� �ٸ� �ؿ����� ������ �ֹ��� �ֽ��ϴ�.\n\n�ؿ����(IDX : " + val_workidx + ") �� �ϰ� �����Ͻðڽ��ϱ�?") != true) {
			return;
		} else {
			// val_workidx = "";
		}
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		if (val_workidx != "") {
			frm.workidx.value = val_workidx;
		}

		frm.submit();
	}
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC" class="a" >
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<tr>
		<td colspan=2>
		<input type=radio name=ctype value="M" <% if ctype="M" then response.write "checked" %> >���� ����(������)
		<input type=radio name=ctype value="J" <% if ctype="J" then response.write "checked" %> >���� ����(�ֹ�����)
		<input type=radio name=ctype value="W" <% if ctype="W" then response.write "checked" %> >Ư�� �Ǹź�

		<input type=checkbox name=onlymifinish <% if onlymifinish="on" then response.write "checked" %> >��ó�� ������
		</td>
	</tr>
	<tr>
		<td >
		������ :
		<% drawSelectBoxOffShopNot000 "shopid",shopid %>
		����� / �Ǹſ� :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class=a>
<form name="buffrm">
<tr>
	<td align=right><input type="text" name="totalsuply" value="" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right"></td>
</tr>
</form>
</table>
<H4>������ ���� �ۼ� �� ������ ���� ���� �� �ּ���.- ���� ���� ����</H4>
<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<form name="frmArr" method=post action="domeaipchulgojungsan.asp">
<input type=hidden name="shopid" value="<%= shopid %>">
<input type=hidden name="workidx" value="">
<% if (ctype="M") or (ctype="J") then %>
	<input type=hidden name="mode" value="chulgo">
	<tr bgcolor="#DDDDFF" align=center>
		<td width=20></td>
		<td width=90>���ó</td>
		<td width=64>����ڵ�</td>
		<td width=64>�ֹ��ڵ�</td>
		<!--
		<td width=30>����</td>
		-->
		<td width=64>������</td>
		<td width=64>�ֹ���/<br>������</td>
		<td width=64>������</td>
		<td width=64>�����</td>
		<td width=64>���ǸŰ�</td>
		<td width=64>�Ѹ��԰�</td>
		<td width=64>�Ѱ��ް�</td>
		<td width=40>��ó��</td>
		<td>���</td>
	</tr>
	<% for i=0 to ofranchulgojungsan.FResultCount-1 %>
	<input type="hidden" name="val_<%= ofranchulgojungsan.FItemList(i).Fid %>" value="<%= ofranchulgojungsan.FItemList(i).Fjumunrealsuplycash %>">
	<%
	ttlsell = ttlsell + ofranchulgojungsan.FItemList(i).Ftotalsellcash
	ttlsuply = ttlsuply + ofranchulgojungsan.FItemList(i).Ftotalsuplycash
	ttlbuy = ttlbuy + ofranchulgojungsan.FItemList(i).Ftotalbuycash
	%>
	<tr bgcolor="#FFFFFF">
		<td ><input type="checkbox" name="check" <% if not IsNULL(ofranchulgojungsan.FItemList(i).Fprecheckidx) then response.write "disabled" %> value="<%= ofranchulgojungsan.FItemList(i).Fid %>" onClick="AnCheckClick(this); reCalcuSum(frmArr);"></td>
		<td ><%= ofranchulgojungsan.FItemList(i).FSocid %></td>
		<td align=center><a href="javascript:PopChulgoSheet('<%= ofranchulgojungsan.FItemList(i).Fid %>')"><%= ofranchulgojungsan.FItemList(i).Fcode %></a></td>
		<td align=center>
			<a href="javascript:PopIpgoSheet('<%= ofranchulgojungsan.FItemList(i).Fbaljuidx %>')">
				<font color="<%= ofranchulgojungsan.FItemList(i).GetOrderStateColor %>"><%= ofranchulgojungsan.FItemList(i).GetOrderStateName %></font>
				<br><%= ofranchulgojungsan.FItemList(i).Fbaljucode %>
			</a>
		</td>
		<!--
		<td align=center><%= ofranchulgojungsan.FItemList(i).Fdivcode %></td>
		-->
		<td align=center><%= ofranchulgojungsan.FItemList(i).Fbaljusegumdate %></td>
		<td align=center><%= Left(ofranchulgojungsan.FItemList(i).FjumunRegDate,10) %><br><%= ofranchulgojungsan.FItemList(i).Fscheduledate %></td>
		<td align=center><%= ofranchulgojungsan.FItemList(i).Fbaljudate %></td>
		<td align=center><%= ofranchulgojungsan.FItemList(i).Fexecutedt %>
			<% if ofranchulgojungsan.FItemList(i).Fexecutedt<>ofranchulgojungsan.FItemList(i).FIpgodate then %>
			<br><font color=red>(<%= ofranchulgojungsan.FItemList(i).FIpgodate %>)</font>
			<% end if %>
		</td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalsellcash,0) %>
			<% if ofranchulgojungsan.FItemList(i).Ftotalsellcash<>ofranchulgojungsan.FItemList(i).Fjumunrealsellcash then %>
			<br><font color=red>(<%= FormatNumber(ofranchulgojungsan.FItemList(i).Fjumunrealsellcash,0) %>)</font>
			<% end if %>
		</td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalbuycash,0) %>
			<% if ofranchulgojungsan.FItemList(i).Ftotalbuycash<>ofranchulgojungsan.FItemList(i).Fjumunrealbuycash then %>
			<br><font color=red>(<%= FormatNumber(ofranchulgojungsan.FItemList(i).Fjumunrealbuycash,0) %>)</font>
			<% end if %>
		</td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %>
			<% if ofranchulgojungsan.FItemList(i).Ftotalsuplycash<>ofranchulgojungsan.FItemList(i).Fjumunrealsuplycash then %>
			<br><font color=red>(<%= FormatNumber(ofranchulgojungsan.FItemList(i).Fjumunrealsuplycash,0) %>)</font>
			<% end if %>
		</td>
		<td align=center>
			<% if not IsNULL(ofranchulgojungsan.FItemList(i).Fprecheckidx) then %>
			<%= ofranchulgojungsan.FItemList(i).Fprecheckmasteridx %>
			<% end if %>
		</td>
		<td>
			<input type="hidden" name="val_workidx_<%= ofranchulgojungsan.FItemList(i).Fid %>" value="<%= ofranchulgojungsan.FItemList(i).Fworkidx %>">
			<% if (ofranchulgojungsan.FItemList(i).Fworkidx <> "") then %>
				�ؿ� IDX : <a href="javascript:PopExportSheet(<%= ofranchulgojungsan.FItemList(i).Fworkidx %>)"><%= ofranchulgojungsan.FItemList(i).Fworkidx %></a>
			<% end if %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlbuy,0) %></td>
		<td align=right><%= formatnumber(ttlsuply,0) %></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" align=center><input type=button value="���� ���� ����" onclick="SaveArr(frmArr)"></td>
	</tr>
<% else %>
	<input type=hidden name="mode" value="witsksell">
	<tr bgcolor="#DDDDFF" align=center>
		<td width=20></td>
		<td width=90>�����</td>
		<td width=90>�ޱ���</td>
		<td width=90>�귣��</td>
		<td width=40>�ѰǼ�</td>
		<td width=80>�Ѹ����</td>
		<td width=80>�Ѹ��԰�<br>(��ü�����)</td>
		<td width=80>�Ѱ��ް�</td>
		<td width=70>������</td>
		<td width=40>��ó��</td>
		<td>���</td>
	</tr>
	<% for i=0 to ofranchulgojungsan.FResultCount-1 %>
	<%
	ttlsell = ttlsell + ofranchulgojungsan.FItemList(i).Ftotsum
	ttlbuy = ttlbuy + ofranchulgojungsan.FItemList(i).Frealjungsansum
	ttlsuply = ttlsuply + 0
	%>
	<tr bgcolor="#FFFFFF">
		<td ><input type="checkbox" name="check" value="<%= ofranchulgojungsan.FItemList(i).Fidx %>" onClick="AnCheckClick(this);"></td>
		<td ><a href="javascript:PopDetail('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= ofranchulgojungsan.FItemList(i).Fshopid %>');"><%= ofranchulgojungsan.FItemList(i).FYYYYMM %></a></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fshopid %></td>
		<td ><a href="javascript:editOffDesinger('<%= ofranchulgojungsan.FItemList(i).Fshopid %>','<%= ofranchulgojungsan.FItemList(i).Fjungsanid %>');"><%= ofranchulgojungsan.FItemList(i).Fjungsanid %></a></td>

		<td align=center><%= ofranchulgojungsan.FItemList(i).Ftotitemcnt %></td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotsum,0) %></td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Frealjungsansum,0) %> </td>
		<td align=right> </td>
		<td></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fprecheckidx %></td>
		<td>
			<input type="hidden" name="val_workidx_<%= ofranchulgojungsan.FItemList(i).Fidx %>" value="">
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlbuy,0) %></td>
		<td align=right><%= formatnumber(ttlsuply,0) %></td>
		<td ></td>
		<td ></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align=center><input type=button value="���� ���� ����" onclick="SaveArr(frmArr)"></td>
	</tr>
<% end if %>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->