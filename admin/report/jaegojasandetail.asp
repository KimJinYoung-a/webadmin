<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jaegocls.asp"-->
<%
dim yyyy1,mm1,designer,mwdiv
yyyy1 = request("yyyy1")
mm1 = request("mm1")
mwdiv = request("mwdiv")
designer = request("designer")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ojaego, yyyymm
yyyymm = yyyy1 + "-" + mm1

set ojaego = new CJaegoEval
ojaego.FRectYYYYMM   = yyyymm
ojaego.FRectMwDiv = mwdiv
ojaego.FRectDesigner = designer
ojaego.GetMonthJeagoDetail


dim totno, totbuy, totsell,i
%>
<h2>������</h2>
�귣�� | ��ǰ�ڵ� | ��ǰ�� | �ɼ� | �Һ��ڰ� | ���԰� | ������ | ������� | �����Ծ�
<br>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frm" method="get" action="">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	���� : <% DrawYMBox yyyy1,mm1 %>
        		�귣�� :	<% drawSelectBoxDesignerwithName "designer", designer %>
        		<input type="radio" name="mwdiv" value="M" <% if mwdiv="M" then response.write "checked" %> >����
        		<input type="radio" name="mwdiv" value="W" <% if mwdiv="W" then response.write "checked" %> >��Ź
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#000000">
    <tr bgcolor="DDDDFF" align="center">
    	<td width="80">�귣��</td>
    	<td width="25">����</td>
    	<td width="50">��ǰ�ڵ�</td>
    	<td>��ǰ��[�ɼ�]</td>
    	<td width="60">����<br>������</td>
    	<td width="60">����Ѿ�<br>(�Һ��ڰ�)</td>
    	<td width="60">����Ѿ�<br>(���԰�)</td>
    	<td width="80">�ݿ������Ѿ�</td>
    	<td width="80">�ݿ������Ѿ�</td>
    	<td width="40">ȸ����</td>
    	<td width="80">3���������Ѿ�</td>
    	<td width="40">ȸ����</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <tr bgcolor="#FFFFFF">
    	<td align="center"><%= ojaego.FItemList(i).Fmakerid %></td>
    	<td></td>
    	<td align="center"><%= ojaego.FItemList(i).Fitemid %></td>
    	<td><%= ojaego.FItemList(i).Fitemname %><br><font color="blue"><%= ojaego.FItemList(i).Fitemoptionname %></font></td>
    	<td align="center"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    	<td>-</td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td align="center">�Ѱ�</td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td align="center" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    </tr>
</table>


<%
set ojaego = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->