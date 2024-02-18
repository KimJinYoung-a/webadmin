<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim yyyy1,mm1,dd1
dim hh1,nn1,ss1
dim shopid,makerid
dim idx

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
hh1 = request("hh1")
nn1 = request("nn1")
ss1 = request("ss1")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2,Cstr(Month(now())))
if (dd1="") then dd1 = Format00(2,Cstr(day(now())))

if (hh1="") then hh1 = "00"
if (nn1="") then nn1 = "00"
if (ss1="") then ss1 = "00"

idx = request("idx")
makerid = request("makerid")
shopid = session("ssBctID")


dim offstock
set offstock = new COffShopDailyStock
offstock.FRectShopId = shopid
offstock.FRectMakerid = makerid


if idx<>"" then
	offstock.FRectIdx = idx
	offstock.GetOneJeagoMaster

	shopid = offstock.FOneItem.FShopid
	makerid = offstock.FOneItem.FMakerid

	offstock.FRectShopID = shopid
	offstock.FRectMakerid = makerid

	offstock.GetDailyStockByInputIdx
else
	offstock.GetDailyStock
end if

dim i, iptot,retot,selltot,currtot

%>
<script language='javascript'>
function searchItems(frm){
	if (frm.shopid.value.length<1){
		alert('��ID�� �����ϼ���.');
		return;
	}

	if (frm.makerid.value.length<1){
		alert('��üID�� �����ϼ���.');
		return;
	}
	frm.submit();
}

function ArrSave(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	var ret;

	upfrm.itemgubunarr.value = "";
	upfrm.shopitemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.realjeagoarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (!IsDigit(frm.realjaego.value)){
				alert('���� ���ڸ� �����մϴ�.');
				frm.realjaego.focus();
				return;
			}

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.shopitemarr.value = upfrm.shopitemarr.value + frm.shopitemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.realjeagoarr.value = upfrm.realjeagoarr.value + frm.realjaego.value + "|";
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		upfrm.submit();
	}
}
</script>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<form name="frm1" method="post" action="realjaegoinput.asp">
<tr>
	<td bgcolor="#DDDDFF" width="100">IDx</td>
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.FIdx %></td>
	<% end if %>
</tr>
<tr>
	<td bgcolor="#DDDDFF">������ID</td>
	<input type=hidden name="shopid" value="<%= shopid %>">
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"><%= shopid %></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fshopid %></td>
	<% end if %>
</tr>
<tr>
	<td bgcolor="#DDDDFF">��üID</td>
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"><%= makerid %></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fmakerid %></td>
	<% end if %>
</tr>
<tr>
	<td bgcolor="#DDDDFF">�ǻ�����Ͻ�</td>
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fjeagodate %></td>
	<% end if %>
</tr>
</table>
<br>
<% if (idx="") then %>
	<table width="800" cellspacing="1" class="a" >
	<tr >
		<td><% DrawOneDateBox yyyy1,mm1,dd1 %>
		&nbsp;
		<input type="text" name="hh1" value="<%= hh1 %>" size=2 maxlength=2>��
		<input type="text" name="nn1" value="<%= nn1 %>" size=2 maxlength=2>��
		<input type="text" name="ss1" value="<%= ss1 %>" size=2 maxlength=2>��
		������ �������</td>
</form>
		<!-- <td align="right"><input type="button" value="��ǰ�˻�" onClick="searchItems(frm1);"></td> -->
	</tr>
	</table>
<% else %>
</form>
<% end if %>

<% if (idx<>"") or ((shopid<>"") and (makerid<>"")) then %>
			<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
			<tr>
			<td colspan="12" align="right" bgcolor="#FFFFFF">
				�ǻ���� ���� ���� �Ͻ� �� "�ǻ���� ����" ��ư�� �����ø� ���� ����˴ϴ�.
			</td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td width="50">�̹���</td>
				<td width="86">���ڵ�</td>
				<td width="100">��ǰ��</td>
				<td width="80">�ɼǸ�</td>
				<td width="80">����<br>�ǻ���</td>
				<td width="50">����<br>�ǻ�</td>
				<td width="50">�԰�</td>
				<td width="50">��ǰ</td>
				<td width="50">�Ǹŷ�</td>
				<td width="50">�������</td>
				<td width="50">�ǻ����</td>
			</tr>
			<% for i=0 to offstock.FresultCount-1 %>
			<%
				iptot = iptot + offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno
				retot = retot + offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno
				selltot = selltot + offstock.FItemList(i).Fsellno
				currtot = currtot + offstock.FItemList(i).Fcurrno
			%>
			<form name="frmBuyPrc_1" >
			<input type="hidden" name="itemgubun" value="<%= offstock.FItemList(i).FItemGubun %>">
			<input type="hidden" name="shopitemid" value="<%= offstock.FItemList(i).FItemId %>">
			<input type="hidden" name="itemoption" value="<%= offstock.FItemList(i).FItemOption %>">
			<tr bgcolor="#FFFFFF">
				<td><img src="<%= offstock.FItemList(i).Fimgsmall %>" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
				<td><%= offstock.FItemList(i).GetBarCode %></td>
				<td><%= offstock.FItemList(i).FItemName %></td>
				<td><%= offstock.FItemList(i).FItemOptionName %></td>
				<td align="center"><%= offstock.FItemList(i).Flastrealdate %></td>
				<td align="center"><%= offstock.FItemList(i).Flastrealno %></td>
				<td align="center"><%= offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno %></td>
				<td align="center"><%= offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno %></td>
				<td align="center"><%= offstock.FItemList(i).Fsellno %></td>
				<% if offstock.FItemList(i).Fcurrno<1 then %>
				<td align="center"><font color="red"><b><%= offstock.FItemList(i).Fcurrno %></font></b></td>
				<% else %>
				<td align="center"><%= offstock.FItemList(i).Fcurrno %></td>
				<% end if %>

				<% if idx<>"" then %>
				<td><input type="text" name="realjaego" value="<%= offstock.FItemList(i).FinputedRealStock %>" size="4" maxlength=8 style="border:1px #999999 solid; text-align=center"></td>
				<% else %>
				<td><input type="text" name="realjaego" value="<%= offstock.FItemList(i).Fcurrno %>" size="4" maxlength=8 style="border:1px #999999 solid; text-align=center"></td>
				<% end if %>
			</tr>
			</form>
			<% next %>
			<tr bgcolor="#FFFFFF">
				<td colspan="5">total</td>
				<td align="center"></td>
				<td align="center"><%= iptot %></td>
				<td align="center"><%= retot %></td>
				<td align="center"><%= selltot %></td>
				<td align="center"><%= currtot %></td>
				<td align="center"></td>
			</tr>
			</table>
	<br>
	<table width="800" cellspacing="1" class="a" >
	<form name="frmArrupdate" method="post" action="shoprealjeago_process.asp">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="designer" value="<%= makerid %>">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="shopitemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="realjeagoarr" value="">

	<tr>
		<td align="right">����ľ� �Ͻ�(��Ȯ�� �Է�) : <% DrawOneDateBox yyyy1,mm1,dd1 %>
		<input type="text" name="hh1" value="<%= hh1 %>" size=2 maxlength=2>��
		<input type="text" name="nn1" value="<%= nn1 %>" size=2 maxlength=2>��
		<input type="text" name="ss1" value="<%= ss1 %>" size=2 maxlength=2>��
		<% if idx<>"" then %>
		<input type="button" value="�ǻ���� ����" onclick="ArrSave()">
		<% else %>
		<input type="button" value="�ǻ���� ����" onclick="ArrSave()">
		<% end if %>
		</td>
	</tr>
	</form>
	</table>
<% end if %>
<%
set offstock = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->