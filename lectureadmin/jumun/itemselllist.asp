<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/jumuncls.asp"-->
<%
'###############################################
' PageName : itemselllist.asp
' Discription : ��ü���� ��ǰ�Ǹ� ���
' History : 2008.07.01 ������ : ���� ��¥�Է��� ���� �ϵ��� ����
'###############################################

dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim orderserial,itemid, itemname
dim datetype
dim oldjumun


orderserial = RequestCheckvar(request("orderserial"),10)
itemid = requestCheckvar(request("itemid"),10)
itemname = requestCheckvar(request("itemname"),100)
searchtype = RequestCheckvar(request("searchtype"),2)
searchrect = RequestCheckvar(request("searchrect"),16)
datetype   = RequestCheckvar(request("datetype"),16)
oldjumun = RequestCheckvar(request("oldjumun"),10)

if (datetype="") then datetype="jumunil"

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

'��ǰ�ڵ� ��ȿ�� �˻�(2008.07.11;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim cknodate,ckdelsearch
cknodate = RequestCheckvar(request("cknodate"),10)
ckdelsearch = RequestCheckvar(request("ckdelsearch"),10)

dim page
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

if ckdelsearch<>"on" then
	ojumun.FRectDelNoSearch="on"
end if


if searchtype="01" then
	ojumun.FRectBuyname = searchrect
elseif searchtype="02" then
	ojumun.FRectReqName = searchrect
elseif searchtype="03" then
	ojumun.FRectUserID = searchrect
elseif searchtype="04" then
	ojumun.FRectIpkumName = searchrect
elseif searchtype="06" then
	ojumun.FRectSubTotalPrice = searchrect
end if

ojumun.FRectItemid = itemid
ojumun.FRectItemName = itemname
ojumun.FRectDesignerID = session("ssBctID")
ojumun.FPageSize = 100
ojumun.FRectOrderSerial = orderserial
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectOldJumun = oldjumun
ojumun.SearchJumunListByDesignerSelllist

dim ix,iy

'response.write ojumun.FRectOrderSerial
'dbget.close()	:	response.End

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
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	��ǰ��ȣ : <input type="text" name="itemid" value="<%= itemid %>" size="11" maxlength="16">
			&nbsp;
			<!--
			��ǰ�� : <input type="text" name="itemname" value="<%= itemname %>" size="11">
			-->
			<br>
			�˻��Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >�ֹ���
			<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >������

			(* �ֱ� 6���� �̳� �ֹ������� �˻��˴ϴ�.)
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
        	�˻���� : �� <b><font color="red"><% = ojumun.FTotalCount %></font></b> ��
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80" align="center">��ǰ�ڵ�</td>
		<td>��ǰ</td>
		<td>�ɼ�</td>
		<td width="80">�ǸŰ�</td>
		<td width="80">���ް�</td>
		<td width="50">����</td>
		<td width="100">���ް��հ�</td>
	</tr>
<% if ojumun.FResultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="12">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FResultCount -1 %>
<%
Dim sumprice,totalsumprice
sumprice = ojumun.FMasterItemList(ix).FBuycash * ojumun.FMasterItemList(ix).FItemNo
%>
	<form name="frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="sitename" value="<%= ojumun.FMasterItemList(ix).FSiteName %>">
	<input type="hidden" name="userid" value="<%= ojumun.FMasterItemList(ix).FUserID %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ojumun.FMasterItemList(ix).FItemID  %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemName %></td>
		<% if (ojumun.FMasterItemList(ix).FItemOptionStr="") then %>
		<td>&nbsp;</td>
		<% else %>
		<td><%= ojumun.FMasterItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0)  %></td>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FBuycash,0)  %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(sumprice,0) %></td>
	</tr>
	</form>
	 <% totalsumprice =  totalsumprice + sumprice %>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" height="25" align="right">���� ������ �հ� �ݾ� : <font color="red"><% =FormatNumber(totalsumprice,0) %></font> ��&nbsp;&nbsp;</td>
	</tr>
<% end if %>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
				<% if ix>ojumun.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if ojumun.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->