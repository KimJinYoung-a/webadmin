<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
rw "��������޴�-�����ڹ��ǿ��"
response.end

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nextdateStr,searchnextdate
dim orderserial,itemid,ojumun
dim topn,shopid,page
dim ckpointsearch,ckipkumdiv4
dim ix,iy,cknodate
dim order_desum
dim rectdispy, rectselly
dim offgubun
dim oldlist

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
shopid = request("shopid")
orderserial = request("orderserial")
itemid = request("itemid")
topn = request("topn")
ckpointsearch = request("ckpointsearch")
cknodate = request("cknodate")
order_desum = request("order_desum")
rectdispy = request("rectdispy")
rectselly = request("rectselly")
offgubun = request("offgubun")
oldlist = request("oldlist")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

topn = request("topn")
if (topn="") then topn=100

set ojumun = new COffShopSellReport

if cknodate="" then
	ojumun.FRectStartDay = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectEndDay = searchnextdate
end if

shopid = "cafe002"
offgubun = "CAF"


ojumun.FRectShopID = shopid
ojumun.FPageSize = topn
ojumun.FCurrPage = page
ojumun.FRectOffgubun = offgubun
ojumun.FRectOldData = oldlist
ojumun.ShopJumunListBybestseller

%>
<script language='javascript'>
function ViewOrderDetail(itemid){
	window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");
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

function ReSearch(ifrm){
	var v = ifrm.topn.value;
	if (!IsDigit(v)){
		alert('���ڸ� �����մϴ�.');
		ifrm.topn.focus();
		return;
	}

	if (v>1000){
		alert('õ�� ���ϸ� �˻������մϴ�.');
		ifrm.topn.focus();
		return;
	}
	ifrm.submit();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3������������
		&nbsp;
		�Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<br>
		�ޱ��� :
		<% if session("ssBctDiv")="101" then %>
			<% 'drawSelectBoxOffShop "shopid",shopid %> 
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %> 
			&nbsp;&nbsp;
		<% else %>
		<% drawSelectBoxOffShopAll "shopid",shopid %> &nbsp;&nbsp;
		<% end if %>
		�˻����� :
		<input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6" >
		<br>
		<input type="radio" name="offgubun" value="" <% if offgubun="" then response.write "checked" %> >����ü
		<input type="radio" name="offgubun" value="OFF" <% if offgubun="OFF" then response.write "checked" %> >����
		<input type="radio" name="offgubun" value="FRN" <% if offgubun="FRN" then response.write "checked" %> >������
		</td>
		<td class="a" align="right">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="1" cellpadding="2" cellspacing="0" class="a">
<tr>
	<td colspan="7" height="25" align="right">�˻���� : �� <font color="red"><% = ojumun.FResultCount %></font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr >
	<td width="100" align="center">��ǰ��ȣ</td>
	<td  align="center">��ǰ</td>
	<td width="80" align="center">�ɼ�</td>
	<td width="100" align="center">����</td>
	<td width="65" align="center">�Ѱ���</td>
	<td width="100" align="center">�հ�ݾ�</td>
</tr>
<% if ojumun.FResultCount<1 then %>
<tr>
	<td colspan="12" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
	<% for ix=0 to ojumun.FResultCount -1 %>
<%
Dim sumprice,totalsumprice
sumprice = ojumun.FItemList(ix).FItemCost * ojumun.FItemList(ix).FItemNo
%>
	<% if ojumun.FItemList(ix).IsAvailJumun then %>
	<tr class="a">
	<% else %>
	<tr class="gray">
	<% end if %>
		<td align="left" height="25"><%= ojumun.FItemList(ix).FItemGubun %>-<%= Format00(6,ojumun.FItemList(ix).FItemID)  %>-<%= ojumun.FItemList(ix).FItemOption %></td>
		<td align="left"><%= ojumun.FItemList(ix).FItemName %></td>
		<% if (ojumun.FItemList(ix).FItemOptionStr="") then %>
			<td align="left">&nbsp;</td>
		<% else %>
			<td align="left"><%= ojumun.FItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FItemCost,0)  %></td>
		<td align="center"><%= ojumun.FItemList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(sumprice,0) %></td>
	</tr>
	 <% totalsumprice =  totalsumprice + sumprice %>
	<% next %>
	<tr>
		<td colspan="7" height="25" align="right">���� ������ �հ� �ݾ� : <font color="red"><% =FormatNumber(totalsumprice,0) %></font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	</tr>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->