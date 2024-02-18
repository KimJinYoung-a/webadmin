<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �����ǰ �� ���������� NO ����¡ ����
' History : 2009.04.07 ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls_datamart.asp"-->
<%
dim shopid , datefg , i ,makerid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim totitemno ,totalsum ,totsuplysum ,totsellsum ,oldlist ,offgubun ,vOffCateCode ,offmduserid
dim vOffMDUserID ,vPurchaseType ,ordertype ,itemid ,itemname ,extbarcode ,reload, buyergubun
dim inc3pl, commCd, arrlist, chkImg
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	menupos = requestCheckVar(request("menupos"),10)
	shopid = requestCheckVar(request("shopid"),32)
	datefg = requestCheckVar(request("datefg"),32)
	makerid = requestCheckVar(request("makerid"),32)
	oldlist = requestCheckVar(request("oldlist"),10)
	offgubun = requestCheckVar(request("offgubun"),32)
	vOffCateCode = requestCheckVar(request("offcatecode"),32)
	vOffMDUserID = requestCheckVar(request("offmduserid"),32)
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	ordertype = requestCheckVar(request("ordertype"),32)
	itemid = requestCheckVar(request("itemid"),10)
	itemname = requestCheckVar(request("itemname"),124)
	extbarcode = requestCheckVar(request("extbarcode"),32)
	reload = requestCheckVar(request("reload"),2)
	buyergubun = requestCheckVar(request("buyergubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    commCd = requestCheckVar(request("commCd"),32)
	chkImg		= requestCheckvar(request("chkImg"),1)

if datefg = "" then datefg = "maechul"
if shopid<>"" then offgubun=""
if ordertype="" then ordertype="totalprice"
if reload <> "on" and offgubun = "" then offgubun = "95"
if chkImg ="" then chkImg = 0

if (yyyy1="") then
	'fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now())))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'C_IS_Maker_Upche = TRUE
'C_IS_SHOP = TRUE

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
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

''��Ÿ�� ������ȸ ����
Dim isFixShopView
IF (session("ssBctID")="doota01") then
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FPageSize = 3000
	ooffsell.FCurrPage = 1
	ooffsell.FRectOldData = oldlist
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.frectdatefg = datefg
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = fromDate
    ooffsell.FRectEndDay = toDate
    ooffsell.FRectDesigner = makerid
	ooffsell.FRectOffgubun = offgubun
	ooffsell.frectoffcatecode = vOffCateCode
	ooffsell.frectoffmduserid = vOffMDUserID
	ooffsell.FRectBrandPurchaseType = vPurchaseType
	ooffsell.FRectOrdertype = ordertype
	ooffsell.FRectitemid = itemid
	ooffsell.FRectitemname = itemname
	ooffsell.FRectextbarcode = extbarcode
	ooffsell.FRectbuyergubun = buyergubun
	ooffsell.FRectInc3pl = inc3pl
	ooffsell.FRectCommCD = commCd
    arrlist = ooffsell.GetDaylySellItemList_getrows

totitemno = 0
totalsum =0
totsuplysum = 0
totsellsum = 0
%>

<script type='text/javascript'>

function frmsubmit(){
	frm.method='get';
	frm.target='';

	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('��ǰ�ڵ�� ���ڸ� �����մϴ�.');
			frm.itemid.focus();
			return;
		}
	}

	frm.submit();
}

function pop_exceldown(){
	frm.action='/admin/offshop/todayselldetail_datamart_excel.asp';
	frm.method='post';
	frm.target='view';
	frm.submit()
	frm.method='get';
	frm.target='';
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reload" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ : <% drawmaechuldatefg "datefg" ,datefg ,""%>
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
					* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
				<br><br>
				* ���屸�� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* ī�װ� : <% SelectBoxBrandCategory "offcatecode", vOffCateCode %>
				&nbsp;&nbsp;
				* ���MD : <% drawSelectBoxCoWorker_OnOff "offmduserid", vOffMDUserID, "off" %>
				&nbsp;&nbsp;
				* �������� : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
				&nbsp;&nbsp;
				* ���Ա��� : <% drawSelectBoxOFFJungsanCommCD "commCd",commCd %>
				<br><br>
				* ��ǰ�ڵ� : <input type="text" name="itemid" value="<%=itemid %>" size=10 maxlength=10>
				&nbsp;&nbsp;
				* ��ǰ�� : <input type="text" name="itemname" value="<%=itemname %>" size=20 maxlength=20>
				&nbsp;&nbsp;
				* ������ڵ� : <input type="text" name="extbarcode" value="<%=extbarcode %>" size=15 maxlength=15>
				&nbsp;&nbsp;
				* ��������: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
				<br><br>
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				<% if C_IS_Maker_Upche then %>
					&nbsp;&nbsp;
					* �귣�� : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					&nbsp;&nbsp;
					* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
			    &nbsp;&nbsp;
			    <input type="checkbox" name="chkImg" value="1" <%if chkImg = 1 then%>checked<%end if%>>��ǰ�̹��� ����
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit();">
	</td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
    	�� ������ �ֹ��� �������� ���� �˴ϴ�.
    </td>
    <td align="right">
    	<input type="button" value="�������" onclick="pop_exceldown();" class="button_s">
    	<% drawordertype "ordertype" ,ordertype ," onchange='frmsubmit();'" ,"I"  %>
    </td>
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		�˻���� : <b><%=ooffsell.FresultCount%></b> ���ִ� 3,000�Ǳ��� ���� �˴ϴ�.(�����ٿ��� 10,000�� ����)
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% IF chkImg = 1 then %>
		<td width=50>�̹���</td>
	<% end if %>
	<td>���ڵ�</td>
	<td>������ڵ�</td>
	<td>�귣��</td>
	<td>��ǰ��(�ɼǸ�)</td>
	<% if (NOT C_InspectorUser) then %>
	<td>�Ǹž�</td>
    <% end if %>
	<td>�����</td>

	<% if not(C_IS_SHOP) then %>
		<td>���Ծ�</td>
	<% end if %>

	<td>�Ǹż���</td>
	<td>���</td>
</tr>
<%
if isarray(arrlist) then

for i=0 to ubound(arrlist,2)

totitemno = totitemno + arrlist(3,i)
totalsum = totalsum + arrlist(0,i)
totsellsum = totsellsum + arrlist(2,i)
totsuplysum = totsuplysum + arrlist(1,i)
%>
<tr bgcolor="#FFFFFF" align="center">
	<% IF chkImg = 1 then %>
		<td>	
			<% if arrlist(17,i)<>"" and not(isnull(arrlist(17,i))) then %>
				<img src="<%= webImgUrl %>/image/small/<%= GetImageSubFolderByItemid(arrlist(11,i)) %>/<%= arrlist(17,i) %>" width="50" height="50" border="0">
			<% else %>
				<img src="<%= webImgUrl %>/offimage/offsmall/i<%= arrlist(10,i) %>/<%= GetImageSubFolderByItemid(arrlist(11,i)) %>/<%= arrlist(18,i) %>" width="50" height="50" border="0">
			<% end if %>
		</td>
	<% end if %>
	<td><%= arrlist(10,i) %><%= BF_GetFormattedItemId(arrlist(11,i)) %><%= arrlist(12,i) %></td>
	<td><%= arrlist(15,i) %></td>
	<td><%= arrlist(13,i) %></td>
	<td align="left">
		<%= arrlist(8,i) %>
		<% if arrlist(9,i) <> "" then %>
			(<%=arrlist(9,i)%>)
		<% end if %>
	</td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(arrlist(2,i),0) %></td>
    <% end if %>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(arrlist(0,i),0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(arrlist(1,i),0) %></td>
	<% end if %>

	<td><%= arrlist(3,i) %></td>
	<td align="center"><%= arrlist(16,i) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<% IF chkImg = 1 then %>
		<td colspan=5><b>�Ѱ�</b></td>
	<% else %>
		<td colspan=4><b>�Ѱ�</b></td>
	<% end if %>
	
	<% if (NOT C_InspectorUser) then %>
	<td align="right"><%= FormatNumber(totsellsum,0) %></td>
    <% end if %>
	<td align="right"><%= FormatNumber(totalsum,0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right"><%= FormatNumber(totsuplysum,0) %></td>
	<% end if %>

	<td><%= FormatNumber(totitemno,0) %></td>
	<td></td>
</tr>
<% else %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
<iframe id="view" name="view" width=0 height=0 frameborder="0" scrolling="no"></iframe>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->