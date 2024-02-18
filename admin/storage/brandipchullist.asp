<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����ۺ����⳻��
' History : 2015.05.27 ���ʻ����� ��
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/AcountItemIpChulCls.asp"-->
<%
dim gubun,designer,itemid, shopid, itemgubun, page, ipchulcode, research
dim IpChulMwgubun, onlineMwDiv, centermwdiv, StockMwDiv, tplgubun, purchasetype
tplgubun = request("tplgubun")
gubun       = RequestCheckVar(request("gubun"),32)
designer    = RequestCheckVar(request("designer"),32)
itemgubun   = RequestCheckVar(request("itemgubun"),2)
itemid      = RequestCheckVar(request("itemid"),9)
shopid      = RequestCheckVar(request("shopid"),32)
page        = RequestCheckVar(request("page"),10)
ipchulcode  = RequestCheckVar(request("ipchulcode"),10)
research  = RequestCheckVar(request("research"),2)
IpChulMwgubun  	= RequestCheckVar(request("IpChulMwgubun"),1)
onlineMwDiv  	= RequestCheckVar(request("onlineMwDiv"),1)
centermwdiv  	= RequestCheckVar(request("centermwdiv"),1)
StockMwDiv  	= RequestCheckVar(request("StockMwDiv"),1)
purchasetype 	= requestCheckVar(request("purchasetype"),3)
''if gubun="" then gubun="I"

if research="" and TPLGubun="" then TPLGubun="3X"


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromDate,toDate

yyyy1 = request("yyyy1")
mm1   = request("mm1")
dd1   = request("dd1")
yyyy2 = request("yyyy2")
mm2   = request("mm2")
dd2   = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oacctipchul
set oacctipchul = new CAcountItemIpChul
oacctipchul.FCurrPage = page
oacctipchul.FPageSize = 1000
oacctipchul.FRectStartday = fromDate
oacctipchul.FRectEndday   = toDate
oacctipchul.FRectGubun   = gubun
oacctipchul.FRectDesigner = designer
oacctipchul.FRectItemGubun = itemgubun
oacctipchul.FRectItemID = itemid
oacctipchul.FRectIpChulCode = ipchulcode
oacctipchul.FtplGubun = tplgubun
oacctipchul.FRectIpChulMwgubun = IpChulMwgubun
oacctipchul.FRectOnlineMwDiv = onlineMwDiv
oacctipchul.FRectCentermwdiv = centermwdiv
oacctipchul.FRectStockMwDiv = StockMwDiv
oacctipchul.FRectBrandPurchaseType = purchasetype

if gubun<>"I" then
	oacctipchul.FRectShopid = shopid
end if

'if (designer<>"") or (itemid<>"") then
    oacctipchul.getIpChulListByBrand
'end if

dim i, sumitemno, sumSellCash, sumBuyCash, sumSuplyCash
sumitemno   = 0
sumSellCash = 0
sumBuyCash  = 0
sumSuplyCash= 0

%>
<script language='javascript'>
function NextPage(ipage){
    document.frm.page.value=ipage;
    document.frm.submit();
}

function jsGoIpChulDetail(iIpChulCode){
    document.frm.ipchulcode.value=iIpChulCode;
    document.frm.submit();
}

function checkDisabled(comp){
    if (comp.value=="I"){
        document.frm.shopid.disabled=true;
    }else{
        document.frm.shopid.disabled=false;
    }
}

function popAssignIpChulMwgubun(didx){
    alert('������ ����');
    <% if (not C_ADMIN_AUTH) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignIpChulMwgubun.asp?didx=" + didx
    var popwin = window.open(iURL,'popAssignIpChulMwgubun','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function jsItemStock(itemgubun, itemid,itemoption){
	var jsItemStock = window.open("/admin/stock/itemcurrentstock.asp?itemgubun="+itemgubun+"&itemid="+itemid+"&itemoption="+itemoption+"&menupos=709","jsItemStock","width=1000 height=600 scrollbars=yes resizable=yes");
	jsItemStock.focus();
}

function popAccStockModiOne(itemgubun,itemid,itemoption){
	var popwin = window.open("/admin/newreport/pop_item_stock_Accsummary_edit.asp?yyyy1=<%=LEFT(now(),4)%>&mm1=<%=MID(now(),6,2)%>&shopid=&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popAccStockModiOne","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popItemIpchulList(makerid, shopid, onlineMwDiv, centermwdiv, StockMwDiv) {
	var frm = document.frm;
	var url = "/admin/storage/itemipchullist.asp?menupos=168";

	url = url + "&designer=" + makerid;
	url = url + "&gubun=<%= gubun %>";
	url = url + "&itemid=<%= itemid %>";
	url = url + "&shopid=" + shopid;
	url = url + "&itemgubun=<%= itemgubun %>";
	url = url + "&ipchulcode=<%= ipchulcode %>";
	url = url + "&IpChulMwgubun=<%= IpChulMwgubun %>";
	url = url + "&onlineMwDiv=" + onlineMwDiv;
	url = url + "&centermwdiv=" + centermwdiv;
	url = url + "&StockMwDiv=" + StockMwDiv;
	url = url + "&tplgubun=<%= tplgubun %>";
	url = url + "&purchasetype=<%= purchasetype %>";

	var popwin = window.open(url,"popItemIpchulList","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		�����ڵ� :
		<input type=text class="text" name="ipchulcode" value="<%= ipchulcode %>" maxlength=10 size=10>
		&nbsp;&nbsp;
		��ǰ�ڵ� :
		<input type=text class="text" name="itemgubun" value="<%= itemgubun %>" maxlength=2 size=2>
		<input type=text class="text" name="itemid" value="<%= itemid %>" maxlength=9 size=8>
		&nbsp;&nbsp;
		�귣��ID :
		<% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;
		    <input type="radio" name="gubun" value="" <% if gubun="" then response.write "checked" %> onClick="checkDisabled(this);">��ü
		    <input type="radio" name="gubun" value="I" <% if gubun="I" then response.write "checked" %> onClick="checkDisabled(this);">�԰�
		    <input type="radio" name="gubun" value="SM" <% if gubun="SM" then response.write "checked" %> onClick="checkDisabled(this);">�����̵�(�������)
			<input type="radio" name="gubun" value="SW" <% if gubun="SW" then response.write "checked" %> onClick="checkDisabled(this);">��������(�������)
			<input type="radio" name="gubun" value="S" <% if gubun="S" then response.write "checked" %> onClick="checkDisabled(this);">�������
		    <input type="radio" name="gubun" value="E" <% if gubun="E" then response.write "checked" %> onClick="checkDisabled(this);">��Ÿ���

		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    �Ⱓ (�������):
		    <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		    ���ó :
		    <% drawSelectBoxChulgo "shopid", shopid %>

		    <% if gubun="I" then %>
			<script language='javascript'>
			document.frm.shopid.disabled=true;
			</script>
			<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    ����ø��Ա��� :
			<select class="select" name="IpChulMwgubun">
				<option value="">����</option>
				<option value="M" <% if (IpChulMwgubun = "M") then %>selected<% end if %> >M</option>
				<option value="F" <% if (IpChulMwgubun = "F") then %>selected<% end if %> >F</option>
				<option value="C" <% if (IpChulMwgubun = "C") then %>selected<% end if %> >C</option>
				<option value="W" <% if (IpChulMwgubun = "W") then %>selected<% end if %> >W</option>
				<option value="X" <% if (IpChulMwgubun = "X") then %>selected<% end if %> >��Ÿ</option>
			</select>
			&nbsp;
			����ON���Ա��� :
			<select class="select" name="onlineMwDiv">
				<option value="">����</option>
				<option value="M" <% if (onlineMwDiv = "M") then %>selected<% end if %> >M</option>
				<option value="W" <% if (onlineMwDiv = "W") then %>selected<% end if %> >W</option>
				<option value="U" <% if (onlineMwDiv = "U") then %>selected<% end if %> >U</option>
				<option value="X" <% if (onlineMwDiv = "X") then %>selected<% end if %> >��Ÿ</option>
			</select>
			&nbsp;
			����OF���͸��Ա��� :
			<select class="select" name="centermwdiv">
				<option value="">����</option>
				<option value="M" <% if (centermwdiv = "M") then %>selected<% end if %> >M</option>
				<option value="W" <% if (centermwdiv = "W") then %>selected<% end if %> >W</option>
				<option value="X" <% if (centermwdiv = "X") then %>selected<% end if %> >��Ÿ</option>
			</select>
			&nbsp;
			�����Ա��� :
			<select class="select" name="StockMwDiv">
				<option value="">����</option>
				<option value="M" <% if (StockMwDiv = "M") then %>selected<% end if %> >M</option>
				<option value="W" <% if (StockMwDiv = "W") then %>selected<% end if %> >W</option>
				<option value="X" <% if (StockMwDiv = "X") then %>selected<% end if %> >��Ÿ</option>
			</select>
			&nbsp;
			3PL���� : <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
			&nbsp;
			�������� : 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchasetype,"" %>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= oacctipchul.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oacctipchul.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
  <td width="80">���ⱸ��</td>
  <% if gubun="I" then %>
  <td width="200">��üID</td>
  <% else %>
  <td width="200">���ó</td>
  <% end if %>
  <td width="200">�귣��ID</td>
  <td width="30">��ǰ<br>����</td>
  <td width="80">�Һ��ڰ�</td>
  <td width="80">���԰�</td>
  <td width="80">���</td>
  <td width="50">����</td>
  <td width="50">����<br>ON����</td>
  <td width="50">����OF<br>���͸���</td>
  <td width="50">���<br>���Ա���</td>
  <td width="50">�������<br>���Ա���</td>
  <td width="50">��������</td>
</tr>
<% if oacctipchul.FResultCount>0 then %>
<% for i=0 to oacctipchul.FResultCount-1 %>
<%
    sumitemno = sumitemno + oacctipchul.FItemList(i).FItemNo
    sumSellCash = sumSellCash + oacctipchul.FItemList(i).FSellCash
    sumBuyCash  = sumbuyCash + Null2Zero(oacctipchul.FItemList(i).FbuyCash)
    sumSuplyCash = sumSuplyCash + oacctipchul.FItemList(i).FsuplyCash
%>
<tr bgcolor="#FFFFFF">
  <td><font color="<%= oacctipchul.FItemList(i).GetDivCodeColor %>"><%= oacctipchul.FItemList(i).GetDivCodeName %></font></td>
  <td><%= oacctipchul.FItemList(i).FSocID %></td>
  <td><a href="javascript:popItemIpchulList('<%= oacctipchul.FItemList(i).Fimakerid %>', '<%= oacctipchul.FItemList(i).FSocID %>', '<%= oacctipchul.FItemList(i).FonlineMwDiv %>', '<%= oacctipchul.FItemList(i).Fcentermwdiv %>', '<%= oacctipchul.FItemList(i).FStockMwDiv %>')"><%= oacctipchul.FItemList(i).Fimakerid %></a></td>
  <td><%= oacctipchul.FItemList(i).FItemgubun %></td>
  <td align="right"><%= FormatNumber(oacctipchul.FItemList(i).FSellCash,0) %></td>
  <% if oacctipchul.FItemList(i).Fipchulflag="I" then %>
    <td align="right"><%= FormatNumber(oacctipchul.FItemList(i).FsuplyCash,0) %></td>
    <td align="right"></td>
  <% else %>
   <td align="right">
    <% if Not isNULL(oacctipchul.FItemList(i).FbuyCash) then %>
    <%= FormatNumber(oacctipchul.FItemList(i).FbuyCash,0) %>
    <% end if %>
    </td>
    <td align="right"><%= FormatNumber(oacctipchul.FItemList(i).FsuplyCash,0) %></td>
  <% end if %>
  <td align="center"><%= oacctipchul.FItemList(i).FItemNo %></td>
  <td align="center"><%= oacctipchul.FItemList(i).FonlineMwDiv %></td>
  <td align="center"><%= oacctipchul.FItemList(i).Fcentermwdiv %></td>
  <td align="center"><%= oacctipchul.FItemList(i).FStockMwDiv %></td>
  <td align="center"><%= oacctipchul.FItemList(i).FStockShopComm_cd %></td>
  <td><%= getBrandPurchaseType(oacctipchul.FItemList(i).Fpurchasetype) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="4"></td>
    <td align="right"><%= FormatNumber(sumSellCash,0) %></td>
    <% if gubun="I" then %>
    <td align="right"><%= FormatNumber(sumSuplyCash,0) %></td>
    <td align="right"></td>
    <% else %>
    <td align="right"><%= FormatNumber(sumBuyCash,0) %></td>
    <td align="right"><%= FormatNumber(sumSuplyCash,0) %></td>
    <% end if %>
	<td align="center"><%= sumitemno %></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
</tr>
<tr height="27" bgcolor="FFFFFF">
	<td colspan="22" align="center">
		<% if oacctipchul.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oacctipchul.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oacctipchul.StarScrollPage to oacctipchul.FScrollCount + oacctipchul.StarScrollPage - 1 %>
			<% if i>oacctipchul.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oacctipchul.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set oacctipchul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
