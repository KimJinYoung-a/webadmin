<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/jungsanCls.asp"-->
<%

dim idx, gubun, rectorder, masteridx
dim yyyymm, i, j
masteridx      = requestCheckvar(request("idx"),10)
gubun   = requestCheckvar(request("gubun"),16)

if gubun="" then gubun="st"

dim sqlStr

dim otpljungsan, otpljungsanmaster, otpljungsanrealdetail, otpljungsangubundetail
set otpljungsanmaster = new CTplJungsan
otpljungsanmaster.FRectIdx = masteridx
otpljungsanmaster.GetTPLJungsanMasterList

if (otpljungsanmaster.FResultCount<1) then
    dbget_TPL.Close : dbget.Close(): response.end
end if


set otpljungsan = new CTplJungsan
otpljungsan.FRectMasterIdx = masteridx
otpljungsan.FRectGubun = gubun
otpljungsan.GetTplJungsanDetailList

yyyymm = otpljungsanmaster.FItemList(0).FYYYYmm


set otpljungsanrealdetail = new CTplJungsan
otpljungsanrealdetail.FRectMasterIdx = masteridx
otpljungsanrealdetail.FRectTplCompanyID = otpljungsanmaster.FItemList(0).Ftplcompanyid

set otpljungsangubundetail = new CTplJungsan

select case gubun
    case "cbm"
        otpljungsanrealdetail.FPageSize = 1000
        otpljungsanrealdetail.GetTplJungsanCbmList
    case else
        otpljungsanrealdetail.FPageSize = 5000
        otpljungsanrealdetail.FRectGubun = gubun
        otpljungsanrealdetail.GetTplJungsanEtcList

        otpljungsangubundetail.FRectGubun = gubun
        otpljungsangubundetail.GetTplJungsanGubunDetailList
end select


dim duplicated

%>
<script>
function addEtcList(iid,igubun){
	window.open('popetclistadd.asp?idx=' + iid + '&gubun=' + igubun,'popetc','width=700, height=150, location=no,menubar=no,resizable=yes,scrollbars=no,status=no,toolbar=no');
}

function DelDetail(frm){
	var ret = confirm('���� ������ ���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function ModiDetail(frm){
	var ret = confirm('���� ������ ���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="idx" value="<%= masteridx %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	��üID : <b><%= otpljungsanmaster.FItemList(0).Ftplcompanyid %></b>
        	&nbsp;
			<input type="radio" name="gubun" value="cbm" <% if gubun="cbm" then response.write "checked" %> > �Ӵ���
			<input type="radio" name="gubun" value="ipchul" <% if gubun="ipchul" then response.write "checked" %> > �������
			<input type="radio" name="gubun" value="etc" <% if gubun="etc" then response.write "checked" %> > ��Ÿ���
        </td>
        <td align="right">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="120">����</td>
        <td width="120">���л�</td>
        <td width="200">����</td>
        <td width="80">�ܰ�</td>
        <td width="80">����</td>
        <td width="80">�ݾ�</td>
        <td width="80">����CBM</td>
        <td width="80">�ݿ�CBM</td>
        <td width="80">���CBM</td>
		<td>�ڸ�Ʈ</td>
    </tr>
    <% for i=0 to otpljungsan.FResultCount-1 %>
    <tr bgcolor="#FFFFFF" align="center" height="25">
      <td ><%= otpljungsan.FItemList(i).Fgubunname %></td>
      <td ><%= otpljungsan.FItemList(i).Fgubundetailname %></td>
      <td ><%= otpljungsan.FItemList(i).Ftypename %></td>
      <td align="right"><%= FormatNumber(otpljungsan.FItemList(i).Funitprice,0) %></td>
      <% if (otpljungsan.FItemList(i).Fgubunname = "�Ӵ��") and (otpljungsan.FItemList(i).Fgubundetailname = "��ǰ����") then %>
      <td align="right"><%= otpljungsan.FItemList(i).Favgcbm %></td>
      <% else %>
      <td align="right"><%= FormatNumber(otpljungsan.FItemList(i).Fitemno,0) %></td>
      <% end if %>
      <td align="right"><%= FormatNumber(otpljungsan.FItemList(i).FtotPrice,0) %></td>
      <% if (otpljungsan.FItemList(i).Fgubunname = "�Ӵ��") and (otpljungsan.FItemList(i).Fgubundetailname = "��ǰ����") then %>
      <td ><%= FormatNumber(otpljungsan.FItemList(i).Fprevcbm, 2) %></td>
      <td ><%= FormatNumber(otpljungsan.FItemList(i).Fcurrcbm, 2) %></td>
      <td ><%= FormatNumber(otpljungsan.FItemList(i).Favgcbm, 2) %></td>
      <% else %>
      <td></td>
      <td></td>
      <td></td>
      <% end if %>
      <td align="left"><%= otpljungsan.FItemList(i).Fcomment %></td>
    </tr>
    <% next %>
</table>

<p />

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<% if gubun="cbm" then %>
            <font color="red"><strong>CBM</strong>(�ִ� 1,000��ǥ��)</font> <%= otpljungsanrealdetail.FResultCount %> ��
			<% elseif gubun="ipchul" then %>
			<font color="red"><strong>�������</strong></font>(�ִ� 5,000��ǥ��) <%= otpljungsanrealdetail.FResultCount %> ��
			<% elseif gubun="witakoffshop" then %>
			<font color="red"><strong>��Ź �������� �Ǹų���</strong></font>(���꿡 ���Ե�)
			<% end if %>
        </td>
        <td align="right">
        	<input type="button" class="button" value="��Ÿ�����߰�" onclick="addEtcList(<%= otpljungsanmaster.FItemList(0).Fidx %>,'<%= gubun %>')">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<% if (gubun = "cbm") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<td width="50">����</td>
<td width="80">��ǰ�ڵ�</td>
<td width="50">�ɼ�</td>
<td width="100">���ڵ�</td>
<td >��ǰ��</td>
<td >�ɼǸ�</td>
<td width="80">����</td>
<td width="80">CBM X(mm)</td>
<td width="80">CBM Y(mm)</td>
<td width="80">CBM Z(mm)</td>
<td width="30">����</td>
<td width="30">����</td>
</tr>
<% for i=0 to otpljungsanrealdetail.FResultCount-1 %>
<form name="frmBuyPrcSell_<%= i %>" method="post" action="dotpljungsan.asp">
<input type="hidden" name="idx" value="<%= otpljungsanrealdetail.FItemList(i).Fidx %>">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fitemgubun %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fitemid %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fitemoption %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fbarcode %></td>
<td ><%= otpljungsanrealdetail.FItemList(i).Fitemname %></td>
<td ><%= otpljungsanrealdetail.FItemList(i).Fitemoptionname %></td>
<td align="center">
    <input type="text" size="3" name="itemno" value="<%= otpljungsanrealdetail.FItemList(i).Fitemno %>" style="text-align:right">
</td>
<td align="center">
    <input type="text" size="3" name="cbmX" value="<%= otpljungsanrealdetail.FItemList(i).FcbmX %>" style="text-align:right">
</td>
<td align="center">
    <input type="text" size="3" name="cbmY" value="<%= otpljungsanrealdetail.FItemList(i).FcbmY %>" style="text-align:right">
</td>
<td align="center">
    <input type="text" size="3" name="cbmZ" value="<%= otpljungsanrealdetail.FItemList(i).FcbmZ %>" style="text-align:right">
</td>
<td ><a href="javascript:DelDetail(frmBuyPrcSell_<%= i %>)">����</a></td>
<td ><a href="javascript:ModiDetail(frmBuyPrcSell_<%= i %>)">����</a></td>
</tr>
</form>
<%
'' ���۱������� �ʰ��� �Ʒ� �ּ�����
if (i mod 1000)=0 then
    response.flush
end if
%>
<% next %>
</table>
<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<td width="120">����</td>
<td width="120">���л�</td>
<td width="200">����</td>
<td width="80">�ܰ�</td>
<td width="80">����</td>
<td width="80">�ݾ�</td>
<td width="120">�����ڵ�</td>
<td>�ڸ�Ʈ</td>
<td width="30">����</td>
<td width="30">����</td>
</tr>
<% for i=0 to otpljungsanrealdetail.FResultCount-1 %>
<form name="frmBuyPrcSell_<%= i %>" method="post" action="dotpljungsan.asp">
<input type="hidden" name="idx" value="<%= otpljungsanrealdetail.FItemList(i).Fidx %>">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fgubunname %></td>
<td align="center"><%= otpljungsanrealdetail.FItemList(i).Fgubundetailname %></td>
<td align="center">
    <select class="select" name="itypename">
    <% for j=0 to otpljungsangubundetail.FResultCount-1 %>
    <% if (otpljungsanrealdetail.FItemList(i).Fgubunname = otpljungsangubundetail.FItemList(j).Fgubunname) and (otpljungsanrealdetail.FItemList(i).Fgubundetailname = otpljungsangubundetail.FItemList(j).Fgubundetailname) then %>
        <option value="<%= otpljungsangubundetail.FItemList(j).Ftypename %>" <%= CHKIIF(otpljungsanrealdetail.FItemList(i).Ftypename = otpljungsangubundetail.FItemList(j).Ftypename, "selected", "") %>><%= otpljungsangubundetail.FItemList(j).Ftypename %></option>
    <% end if %>
    <% next %>
    </select>
</td>
<td align="right">
    <input type="text" size="3" name="unitprice" value="<%= otpljungsanrealdetail.FItemList(i).Funitprice %>" style="text-align:right">
</td>
<td align="right">
    <input type="text" size="3" name="itemno" value="<%= otpljungsanrealdetail.FItemList(i).Fitemno %>" style="text-align:right">
</td>
<td align="right"><%= FormatNumber(otpljungsanrealdetail.FItemList(i).FtotPrice, 0) %></td>
<td align="center">
    <%= otpljungsanrealdetail.FItemList(i).Fmastercode %>
</td>
<td align="left">
    <%= otpljungsanrealdetail.FItemList(i).Fcomment %>
</td>
<td ><a href="javascript:DelDetail(frmBuyPrcSell_<%= i %>)">����</a></td>
<td ><a href="javascript:ModiDetail(frmBuyPrcSell_<%= i %>)">����</a></td>
</tr>
</form>
<%
'' ���۱������� �ʰ��� �Ʒ� �ּ�����
if (i mod 1000)=0 then
    response.flush
end if
%>
<% next %>
</table>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->
