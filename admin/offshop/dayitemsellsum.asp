<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
' History : 2011.12.27 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim yyyy1, mm1, dd1, yyyy2, mm2,dd2 , itemgubun , itemid , itemname ,page ,datefg ,ooffsell,i
dim yyyymmdd1,yyymmdd2 ,shopid, makerid ,offgubun , oldlist ,nextdateStr,searchnextdate
Dim totitemno , totsellprice , totrealsellprice ,totsuplyprice, extbarcode, inc3pl
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	shopid = requestCheckVar(request("shopid"),32)
	makerid = requestCheckVar(request("makerid"),32)
	offgubun = requestCheckVar(request("offgubun"),10)
	oldlist = requestCheckVar(request("oldlist"),10)
	itemgubun = requestCheckVar(request("itemgubun"),2)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	page    = requestCheckVar(request("page"),10)
	datefg = requestCheckVar(request("datefg"),32)
	extbarcode = requestCheckVar(request("extbarcode"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"
if page="" then page="1"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

'C_IS_SHOP = TRUE
'C_IS_Maker_Upche = TRUE

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"GREENBEE_1"
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

if shopid<>"" then offgubun=""

''��Ÿ�� ������ȸ ����
Dim isFixShopView
IF (session("ssBctID")="doota01") then
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

set ooffsell = new COffShopSellReport
	ooffsell.FRectStartDay = yyyy1 + "-" + mm1 + "-" + dd1
	ooffsell.FRectEndDay = searchnextdate
	ooffsell.frectitemname = itemname
	ooffsell.frectdatefg = datefg
	ooffsell.FPageSize = 500
	ooffsell.FCurrPage = page
	ooffsell.FRectShopID = shopid
	ooffsell.frectitemid = itemid
	ooffsell.FRectDesigner = makerid
	ooffsell.FRectOffgubun = offgubun
	ooffsell.FRectOldData = oldlist
	ooffsell.frectitemgubun = itemgubun
	ooffsell.frectextbarcode = extbarcode
	ooffsell.FRectInc3pl = inc3pl
	ooffsell.getdayitemsum

totitemno = 0
totsellprice = 0
totrealsellprice= 0
totsuplyprice = 0
%>

<script language="javascript">

	function formsubmit(page){

		if(frm.itemid.value!=''){
			if (!IsDouble(frm.itemid.value)){
				alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
				frm.itemid.focus();
				return;
			}
		}

		frm.page.value=page;
		frm.submit();
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ : <% drawmaechul_datefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3������
		&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>
			<% if (not C_IS_OWN_SHOP and shopid <> "") or (isFixShopView) then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			<% if not(C_IS_Maker_Upche) then %>
				* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% else %>
				* ���� : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," ","" %>
			<% end if %>
		<% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="formsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �귣�� :
		<% if C_IS_Maker_Upche then %>
			<%= makerid %>
		<% else %>
			<% drawSelectBoxDesignerwithName "makerid",makerid %>
		<% end if %>
		&nbsp;&nbsp;
		* ��ǰ�ڵ� : <input type="text" name="itemid" value="<%=itemid%>" size=10>
		&nbsp;&nbsp;
		* ��ǰ�� : <input type="text" name="itemname" value="<%=itemname%>">
		&nbsp;&nbsp;
		* �����ڵ� : <input type="text" name="extbarcode" value="<%=extbarcode%>" size=14 maxlength=14>
		<p>
		* ���� ���� : <% Call DrawShopDivCombo("offgubun",offgubun) %>
        &nbsp;&nbsp;
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>

		<% if not(C_IS_Maker_Upche) then %>
			&nbsp;&nbsp;
			* ��ǰ���� :
			<% drawSelectBoxItemGubun "itemgubun" , itemgubun %>
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">
    <tr valign="bottom">
        <td align="left">
        	�� ������ �ֹ��� �������� ���� �˴ϴ�.
	    </td>
	    <td align="right">
        </td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ooffsell.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ooffsell.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��¥</td>
	<td>��ǰ��ȣ</td>
	<td>��ǰ��(�ɼǸ�)</td>
	<td>�귣��</td>
	<td>�ǸŰ�</td>
	<td>�����</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>���԰�</td>
	<% end if %>

	<td>�Ǹż���</td>
	<td>���</td>
</tr>
<%
if ooffsell.FresultCount > 0 then

for i=0 to ooffsell.FresultCount-1

totitemno = totitemno + ooffsell.FItemList(i).FItemNo
totsellprice = totsellprice + (ooffsell.FItemList(i).FItemCost * ooffsell.FItemList(i).FItemNo)
totrealsellprice = totrealsellprice + (ooffsell.FItemList(i).frealsellprice * ooffsell.FItemList(i).FItemNo)
totsuplyprice = totsuplyprice + (ooffsell.FItemList(i).fsuplyprice * ooffsell.FItemList(i).FItemNo)
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<td><%= getweekendcolor(ooffsell.FItemList(i).fIXyyyymmdd) %></td>
	<td><%= ooffsell.FItemList(i).FItemGubun %>-<%= CHKIIF(ooffsell.FItemList(i).FItemID>=1000000,Format00(8,ooffsell.FItemList(i).FItemID),Format00(6,ooffsell.FItemList(i).FItemID))  %>-<%= ooffsell.FItemList(i).FItemOption %></td>
	<td align="left">
		<%= ooffsell.FItemList(i).FItemName %>
		<% if (ooffsell.FItemList(i).FItemOptionStr<>"") then %>
		(<%= ooffsell.FItemList(i).FItemOptionStr %>)
		<% end if %>
	</td>
	<td ><%= ooffsell.FItemList(i).FMakerid %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FItemCost,0)  %></td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).frealsellprice,0)  %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).fsuplyprice,0)  %></td>
	<% end if %>

	<td><%= ooffsell.FItemList(i).FItemNo %></td>
	<td></td>
</tr>
<% next %>

<tr bgcolor="#FFFFFF" align="center">
	<td colspan=4>�հ�</td>
	<td align="right"><%=FormatNumber(totsellprice,0) %></td>
	<td align="right"><%=FormatNumber(totrealsellprice,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%=FormatNumber(totsuplyprice,0) %></td>
	<% end if %>

	<td><%=FormatNumber(totitemno,0) %></td>
	<td></td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if ooffsell.HasPreScroll then %>
			<span class="list_link"><a href="javascript:formsubmit('<%= ooffsell.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ooffsell.StartScrollPage to ooffsell.StartScrollPage + ooffsell.FScrollCount - 1 %>
			<% if (i > ooffsell.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ooffsell.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:formsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ooffsell.HasNextScroll then %>
			<span class="list_link"><a href="javascript:formsubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="20">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
set ooffsell = Nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->