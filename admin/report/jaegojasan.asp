<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jaegocls.asp"-->
<H3>������ - ������</H3>
<%
dim yyyy1,mm1,designer,mwgubun,isusing
designer = request("designer")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
isusing = request("isusing")
mwgubun = request("mwgubun")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ojaego, yyyymm, enddate, pre3month
yyyymm = yyyy1 + "-" + mm1
enddate = dateserial(yyyy1,mm1+1,1)
pre3month = dateserial(yyyy1,mm1-2,1)

set ojaego = new CJaegoEval
ojaego.FRectYYYYMM   = yyyymm
ojaego.FRectIsusing = isusing
'ojaego.FRectDesigner = designer
ojaego.GetMonthJeagoSum

dim ojaegomaker
set ojaegomaker = new CJaegoEval

if mwgubun<>"" then
	ojaegomaker.FRectYYYYMM = yyyymm
	ojaegomaker.FRectStartDate = yyyymm + "-01"
	ojaegomaker.FRectEndDate = CStr(enddate)
	ojaegomaker.FRect3MonthStartDate = CStr(pre3month)
	ojaegomaker.FRectMwDiv = mwgubun
	''ojaegomaker.GetMonthJeagoSumByMaker
end if

dim totno, totbuy, totsell,i
dim totonlinemeaip, totofflinemeaip, totoffchulgobuycash
dim totoffchulgosuplycash, totFMonthMeachulSum, tot3MonthMeachulSum, totoff3monthchulgosuplycash
%>
<script language='javascript'>
function popStockJasan(mwdiv,yyyy1,mm1,designer){
	var popwin = window.open("jaegojasandetail.asp?mwdiv=" + mwdiv + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&designer=" + designer,"stockdetail","width=1000,height=620,scrollbars=yes, resizable=yes");
	popwin.focus();
}
</script>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>��������ڻ� �� ȸ����</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br>���������� ���� ����ڻ� �� �귣�庰 ȸ���� �����Դϴ�.
			<br>
			<br>����� ��ǰ�� �⺻��������....
			<br>�����ϴ� �귣�� ��ǥ��
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	�˻� : <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
	        	&nbsp;&nbsp;&nbsp;
	        	��뱸��:
	        	<input type="radio" name="isusing" value="">��ü
	        	<input type="radio" name="isusing" value="Y">�����
	        	<input type="radio" name="isusing" value="N">������
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="100">���Ա���</td>
    	<td width="100">��������</td>
    	<td width="100">�Һ��ڰ�</td>
    	<td width="100">��ո���</td>
    	<td width="100">���԰�</td>
    	<td>���</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="?menupos=<%= menupos %>&mwgubun=<%= ojaego.FItemList(i).FMaeIpGubun %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td><%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�Ѱ�</td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>

    	<td></td>
    </tr>
</table>

<% if mwgubun="M" then %>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
            <br>
            * ���� ���� - �귣�庰<br>
            * ���� ���� - ���� Ư�� ���о��� �� ����
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<% elseif mwgubun="W" then %>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
            <br>
            * Ư�� ���� - �귣�庰
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<% end if %>

<%
totno = 0
totbuy = 0
totsell = 0
%>

<% if mwgubun<>"" then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td rowspan=2>�귣��</td>
    	<td rowspan=2>�����</td>
        <td rowspan=2>����Ѿ�<br>(���԰�)</td>
    <!--	<td rowspan=2>����<br>(�Һ��ڰ�)</td>     -->
    	<td colspan=2><%= mm1 %>�� �Ѹ��Ծ�</td>
    	<td rowspan=2><%= mm1 %>�� ��������<br>���Ծ�</td>

    	<td colspan=3><%= mm1 %>�� ȸ����

    	<td colspan=3>3���� ȸ����
    </tr>
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="100"><%= mm1 %>�� ���Ծ�</td>
    	<td width="100"><%= mm1 %>�� off��������</td>

    	<td width="100"><%= mm1 %>�� �¶���<br>�����</td>
    	<td width="100"><%= mm1 %>�� ��������<br>����(���)��</td>
    	<td width="100"><%= mm1 %>��ȸ����</td>

    	<td width="100">3���� �¶���<br>�����</td>
    	<td width="100">3���� ��������<br>����(���)��</td>
    	<td width="100">3����ȸ����</td>
    </tr>
    <% for i=0 to ojaegomaker.FResultCount -1 %>
    <%
        totno   = totno + ojaegomaker.FItemList(i).FTotCount
        totbuy  = totbuy + ojaegomaker.FItemList(i).FTotBuySum
        totsell = totsell + ojaegomaker.FItemList(i).FTotSellSum

        totonlinemeaip = totonlinemeaip + ojaegomaker.FItemList(i).Fonlinemeaip
        totofflinemeaip = totofflinemeaip + ojaegomaker.FItemList(i).Fofflinemeaip
        totoffchulgobuycash = totoffchulgobuycash + ojaegomaker.FItemList(i).Foffchulgobuycash
        totoffchulgosuplycash = totoffchulgosuplycash + ojaegomaker.FItemList(i).Foffchulgosuplycash
        totoff3monthchulgosuplycash = totoff3monthchulgosuplycash + ojaegomaker.FItemList(i).Foff3monthchulgosuplycash
        totFMonthMeachulSum = totFMonthMeachulSum + ojaegomaker.FItemList(i).FMonthMeachulSum
        tot3MonthMeachulSum = tot3MonthMeachulSum + ojaegomaker.FItemList(i).F3MonthMeachulSum

    %>
    <tr bgcolor="#FFFFFF">
    	<td><a href="javascript:popStockJasan('<%= mwgubun %>','<%= yyyy1 %>','<%= mm1 %>','<%= ojaegomaker.FItemList(i).Fmakerid %>');"><%= ojaegomaker.FItemList(i).Fmakerid %></a></td>
    	<td align="center"><%= FormatNumber(ojaegomaker.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).FTotBuySum,0) %></td>
    <!--	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).FTotSellSum,0) %></td>  -->
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Fonlinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Fofflinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Foffchulgobuycash,0) %></td>

    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).FMonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Foffchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if ojaegomaker.FItemList(i).FTotBuySum<>0 then %>
    		<%= CLng((ojaegomaker.FItemList(i).Foffchulgosuplycash+ojaegomaker.FItemList(i).FMonthMeachulSum)/ojaegomaker.FItemList(i).FTotBuySum*100)/100 %>
    	<% end if %>
    	</td>

    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).F3MonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Foff3monthchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if ojaegomaker.FItemList(i).FTotBuySum<>0 then %>
    		<%= CLng((ojaegomaker.FItemList(i).Foff3monthchulgosuplycash+ojaegomaker.FItemList(i).F3MonthMeachulSum)/ojaegomaker.FItemList(i).FTotBuySum*100)/100 %>
    	<% end if %>
    	</td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td></td>
    	<td align="center"><%= FormatNumber(totno,0) %></td>
    	<td align="right"><%= FormatNumber(totbuy,0) %></td>
    <!--	<td align="right"><%= FormatNumber(totsell,0) %></td>   -->
    	<td align="right"><%= FormatNumber(totonlinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(totofflinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(totoffchulgobuycash,0) %></td>

    	<td align="right"><%= FormatNumber(totFMonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(totoffchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if totbuy<>0 then %>
    		<%= CLng((totoffchulgosuplycash+totFMonthMeachulSum)/totbuy*100)/100 %>
    	<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(tot3MonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(totoff3monthchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if totbuy<>0 then %>
    		<%= CLng((totoff3monthchulgosuplycash+tot3MonthMeachulSum)/totbuy*100)/100 %>
    	<% end if %>
    	</td>
    </tr>
</table>
<% end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set ojaegomaker = Nothing
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->