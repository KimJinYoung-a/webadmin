<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �¶��� �������
' History : 2007.12.06 �ѿ�� ����
'			2011.05.18 ������ ����(�Һ�, ���αݾ�, ��ǰ�������׵� �߰�/ ���� = ���Ծ�/�ǰ��������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<%
dim dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc, exceptChangeOrd, research, grpTp
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim i ,defaultdate,defaultdate1 , olddata
dim channelDiv, inc3pl
	ipkumdatesucc = request("ipkumdatesucc")
	olddata = request("olddata")
	sitename = request("sitenamebox")
	accountdiv = request("accountdiv")
	bancancle = request("bancancle")
	if bancancle = "" then bancancle = "1"
	datecancle = request("datecancle")
	dateview1 = request("dateview1")
	if dateview1 = "" then dateview1 = "yes"
	defaultdate1 = dateadd("d",-60,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 60�������� �˻�
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = mid(defaultdate1,6,2)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = right(defaultdate1,2)
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	mm2 = request("mm2")
	if mm2 = "" then
		mm2 = month(now)
	else
		if Len(mm2) = 2 then
			mm2 = request("mm2")
		else
			mm2 = "0"&request("mm2")
		end if
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)

	channelDiv = request("channelDiv")

	research   = request("research")
	exceptChangeOrd = request("exceptChangeOrd")
    grpTp = request("grpTp")
    inc3pl = request("inc3pl")
	if (research="") then exceptChangeOrd="on"
	if (grpTp="") then grpTp="d"

dim Omaechul_list
set Omaechul_list = new Cmaechul_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
	Omaechul_list.FRectEndDate = yyyy2 & "-" & mm2 & "-" & dd2
	Omaechul_list.frectdatecancle = datecancle
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc
	Omaechul_list.fRectChannelDiv=channelDiv
	Omaechul_list.fRectexceptChangeOrd=exceptChangeOrd
	Omaechul_list.FRectGroupType=grpTp
	Omaechul_list.FRectInc3pl = inc3pl  ''2013/12/02 �߰�
	Omaechul_list.fmaechul_list()

if olddata = "no" then
	dim Omaechul_list_old
'	set Omaechul_list_old = new Cmaechul_list
'		Omaechul_list_old.FRectStartdate = (yyyy1-1) & "-" & mm1 & "-" & dd1
'		Omaechul_list_old.FRectEndDate = (yyyy2-1) & "-" & mm2 & "-" & dd2
'		Omaechul_list_old.frectdatecancle = datecancle
'		Omaechul_list_old.frectbancancle = bancancle
'		Omaechul_list_old.frectaccountdiv = accountdiv
'		Omaechul_list_old.frectsitename = sitename
'		Omaechul_list_old.fRectChannelDiv=channelDiv
'		Omaechul_list_old.fmaechul_list()
end if

''����Ʈ����
Sub Drawsitename(selectboxname, sitename)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�._selectboxname�� sub���������� ����
	dim userquery, tem_str

	response.write "<select name='" & selectboxname & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if sitename ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">��ü</option>"								'�����̶� �ܾ ��������.

	'����� �˻� �ɼ� ���� DB���� ��������
	userquery = " select id from [db_partner].[dbo].tbl_partner"
	userquery = userquery + " where 1=1"
	userquery = userquery + " and id <> ''"
	userquery = userquery + " and id is not null"
	userquery = userquery + " and userdiv= '999'"
	userquery = userquery + " group by id"

	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("id")) then 	'�˻��� �̸��� db�� ����� �̸��� ���ؼ� �´ٸ�, //
				tem_str = " selected"								'// �˻���� ����
			end if

			response.write "<option value='" & rsget("id") & "' " & tem_str & ">" & rsget("id") & "</option>"
			tem_str = ""				'rsget�� id�� �����ϰ� �˻��� ������ ����
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub

Dim vParameter
	vParameter = "yyyy1="&yyyy1&"&yyyy2="&yyyy2&"&datecancle="&datecancle&"&bancancle="&bancancle&"&accountdiv="&accountdiv&"&sitename="&sitename&"&dateview1="&dateview1&"&ipkumdatesucc="&ipkumdatesucc&""
%>

<script language="javascript" src="/admin/maechul/daumchart/FusionCharts.js"></script>		<!-- �׷����� ���� �ڹٽ�ũ��Ʈ����-->
<script language="javascript">

function submit()
{
	frm.submit();
}

<!--���� ���� �󼼺��� ����-->
function monthsum(yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var monthsum = window.open('/admin/maechul/maechul_month_sum.asp?yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'monthsum','width=1024,height=768,scrollbars=yes,resizable=yes');
	monthsum.focus();
}
<!--���� ���� �󼼺��� ��-->

<!--���� ���� �󼼺��� ����-->
function weeksum(yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var weeksum = window.open('/admin/maechul/maechul_week_sum.asp?yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'weeksum','width=1024,height=768,scrollbars=yes,resizable=yes');
	weeksum.focus();
}
<!--���� ���� �󼼺��� ��-->

<!--������� ����-->
function excelprint(olddata,yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var excelprint = window.open('/admin/maechul/maechul_sum_excel.asp?olddata='+olddata+'&yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'excelprint','width=1024,height=768,scrollbars=yes,resizable=yes');
	excelprint.focus();
}
<!--���� ���  ��-->

function goOpenGraph()
{
	var graph = window.open('pop_graph.asp','graph','width=1024, height=768, scrollbars=yes, resizable=yes');
	graph.focus();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
            * �Ⱓ :
			<select name="dateview1" class="select">
				<option value="yes" <%=CHKIIF(dateview1="yes","selected","")%>>�ֹ���</option>
				<option value="no" <%=CHKIIF(dateview1="no","selected","")%>>������</option>
			</select>
			<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			<input type="radio" name="grpTp" value="d" <%=CHKIIF(grpTp="d","checked","") %> >�Ϻ�
			<input type="radio" name="grpTp" value="m" <%=CHKIIF(grpTp="m","checked","") %> >����

		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

        	<!--<input type=checkbox name="datecancle" value="on" <% if datecancle="on" then  response.write "checked" %>>��ҰǸ�-->
        	* ����Ʈ���� : <% Drawsitename "sitenamebox",sitename %>
			* �ֹ����� :
			<select name="bancancle" class="select">
				<option value="1" <%=CHKIIF(bancancle="1","selected","")%>>��ǰ����</option>
				<option value="3" <%=CHKIIF(bancancle="3","selected","")%>>��ǰ����</option>
				<option value="2" <%=CHKIIF(bancancle="2","selected","")%>>��ǰ�Ǹ�</option>
			</select>
        	* �������� <select name="accountdiv">
        		<option value="" <% if accountdiv = "" then response.write "selected" %>>��ü</option>
        		<option value="7" <% if accountdiv = "7" then response.write "selected" %>>������</option>
				<option value="14" <% if accountdiv = "14" then response.write "selected" %>>����������</option>
        		<option value="20" <% if accountdiv = "20" then response.write "selected" %>>�ǽð�</option>
        		<option value="50" <% if accountdiv = "50" then response.write "selected" %>>�ܺθ�</option>
        		<option value="80" <% if accountdiv = "80" then response.write "selected" %>>�ÿ�</option>
        		<option value="100" <% if accountdiv = "100" then response.write "selected" %>>�ſ�ī��</option>
        	</select>
        	* ä�α���
        	<select name="channelDiv">
	        	<option value="" <%=CHKIIF(channelDiv="","selected","") %> >��ü</option>
	        	<option value="w" <%=CHKIIF(channelDiv="w","selected","") %> >��</option>
	        	<option value="j" <%=CHKIIF(channelDiv="j","selected","") %> >����</option>
	        	<option value="m" <%=CHKIIF(channelDiv="m","selected","") %> >�������</option>
        	</select>
        	<input type=checkbox name="exceptChangeOrd" value="on" <% if exceptChangeOrd="on" then  response.write "checked" %>>��ȯ�ֹ�����
        	<input type=checkbox name="ipkumdatesucc" value="on" <% if ipkumdatesucc="on" then  response.write "checked" %>>�̰���������
            &nbsp;
            <b>* ����ó����</b>
        	<% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
		</td>
	</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left" style="padding:10 0 10 0">
		&nbsp;
		<!-- ������
			<input type="button" class="button" value="�������" onclick="javascript:excelprint('<%=olddata%>','<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
	    -->
		</td>
		<td align="right">
		<% if (NOT C_InspectorUser) then %>
			<input type="button" class="button" value="�׷��� ���" onclick="javascript:goOpenGraph();">
		<% end if %>
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ���� �ֺ� ������� �󼼳��� ���� ����-->
<!-- radio��ư���� ����
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor=FFFFFF>
		<td align="left" style="padding:3 0 3 0">
			&nbsp;&nbsp;<a href="javascript:monthsum('<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
			���� ������� �󼼳��� ���� [Ŭ��]</a>
		</td>
		<td align="right">
			<a href="javascript:weeksum('<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
			�ֺ� ������� �󼼳��� ���� [Ŭ��]</a>&nbsp;&nbsp;
		</td>
	</tr>
</table>
-->
<!-- ���� �ֺ� ������� �󼼳��� ���� ��-->

<!-- ����Ʈ ���� -->
<%
dim totalsum_totalsum, totalcount_totalsum, subtotalprice_totalsum, totalbuysum_totalsum, spendScoupon_totalsum, spendMileage_totalsum
dim discountEtc_totalsum, sumpaymentEtc_totalsum, tendeliverBuysum_totalsum, tendeliverCount_totalsum, sunsuik_totalsum, magin_totalsum
Dim TTLtotalorgitemcostsum,TTLtotalOrgDlvPay,TTLtotalitemcostcouponNotApplied,TTLtotalCouponNotAppliedDlvPay
Dim TTLtotalitemcostsum,TTLtotalDlvPay,TTLtotalreducedDlvPay,TTLupchepartDeliverBuySum
%>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if dateview1 = "yes" then %>
		<td align="center" width="70" rowspan="2">�ֹ���</td>
	<% elseif dateview1 = "no" then %>
		<td align="center" width="70" rowspan="2">�Ա���</td>
	<% end if %>
	<% if datecancle <> "" then %>
		<td align="center" width="70" rowspan="2">�����</td>
	<% end if %>
    <td align="center" width="50" rowspan="2">���ֹ�<br>�Ǽ�</td>
<% if (NOT C_InspectorUser) THEN %>
    <td align="center" colspan="2">�Һ��ڰ�<br>A</td>
    <td align="center" colspan="2">���αݾ�<br>B</td>
    <td align="center" colspan="2">�ǸŰ�(���ΰ�)<br>C=A-B</td>
    <td align="center" colspan="2">��ǰ��������<br>D</td>
    <td align="center" colspan="2">�����Ѿ�<br>E=C-D</td>
    <td align="center" colspan="2">���ʽ���������<br>F</td>
	<td align="center" width="70" rowspan="2">��Ÿ����<br>H</td>
<% end if %>
	<td align="center" width="70" rowspan="2">�����<br>E-F-H</td>
	<td align="center" width="70" rowspan="2">���ϸ���<br>G</td>
	<td align="center" width="70" rowspan="2">��ġ��<br>���<br>M</td>
	<td align="center" width="70" rowspan="2"><strong>�����Ѿ�</strong><br>I=E-G</td>
	<td align="center" width="70" rowspan="2">���԰�<br>(��ǰ����)<br>J</td>
	<td align="center" colspan="2">��ۺ��<br>K</td>
	<td align="center" rowspan="2">����<br>L=(J+K)/I</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if (NOT C_InspectorUser) THEN %>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
    <td>��ǰ</td>
    <td>��ۺ�</td>
<% end if %>
    <td>�ٹ�</td>
    <td>����</td>
</tr>

<% for i = 0 to Omaechul_list.ftotalcount -1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td align="center">
		<% if (grpTp="d") and right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "�����" then %>
			<font color="blue"><%= Omaechul_list.flist(i).forderdate %></font>
		<% elseif (grpTp="d") and right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "�Ͽ���" then %>
			<font color="red"><%= Omaechul_list.flist(i).forderdate %></font>
		<% else %>
			<%= Omaechul_list.flist(i).forderdate %>
		<% end if %>
	</td>
    <td align="center"><%= Omaechul_list.flist(i).ftotalcount %></td>
<% if (NOT C_InspectorUser) THEN %>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalOrgDlvPay) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum-Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalOrgDlvPay-Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied-Omaechul_list.flist(i).ftotalitemcostsum) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay-Omaechul_list.flist(i).ftotalDlvPay) %></td>

    <% if IsNULL(Omaechul_list.flist(i).ftotalitemcostsum) then %>
    	<td align="right" colspan="2" bgcolor="#9DCFFF"><%= CurrFormat(Omaechul_list.flist(i).ftotalsum) %></td>
    <% else %>
	    <td align="right" bgcolor="#9DCFFF"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostsum) %></td>
	    <td align="right" bgcolor="#9DCFFF"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalDlvPay) %></td>
    <% end if %>

    <% if IsNULL(Omaechul_list.flist(i).ftotalreducedDlvPay) then %>
    	<td align="right" colspan="2"><%= CurrFormat(Omaechul_list.flist(i).fspendScoupon) %></td>
    <% else %>
    	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fspendScoupon-(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay)) %></td>
    	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay) %></td>
    <% end if %>
	<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdiscountEtc) %></td>
<% end if %>
    <td align="right" bgcolor="#E6B9B8">
    	<%= CurrFormat((Omaechul_list.flist(i).ftotalitemcostsum+Omaechul_list.flist(i).ftotalDlvPay)-((Omaechul_list.flist(i).fspendScoupon-(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay))+(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay))-(Omaechul_list.flist(i).fdiscountEtc)) %>
    </td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendMileage) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsumpaymentetc) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsubtotalprice) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalbuysum) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftendeliverBuysum) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fupchepartDeliverBuySum) %></td>
    <td align="center"><%= Omaechul_list.flist(i).fmagin %>%</td>
</tr>
    <%
    totalcount_totalsum = totalcount_totalsum + Omaechul_list.flist(i).ftotalcount
    TTLtotalorgitemcostsum = TTLtotalorgitemcostsum + Omaechul_list.flist(i).ftotalorgitemcostsum
    TTLtotalOrgDlvPay      = TTLtotalOrgDlvPay      + Omaechul_list.flist(i).ftotalOrgDlvPay
    TTLtotalitemcostcouponNotApplied = TTLtotalitemcostcouponNotApplied + Omaechul_list.flist(i).ftotalitemcostcouponNotApplied
    TTLtotalCouponNotAppliedDlvPay = TTLtotalCouponNotAppliedDlvPay + Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay
    TTLtotalitemcostsum = TTLtotalitemcostsum + Omaechul_list.flist(i).ftotalitemcostsum
    TTLtotalDlvPay = TTLtotalDlvPay + Omaechul_list.flist(i).ftotalDlvPay

    TTLtotalreducedDlvPay = TTLtotalreducedDlvPay + Omaechul_list.flist(i).ftotalreducedDlvPay

    totalsum_totalsum = totalsum_totalsum + Omaechul_list.flist(i).ftotalsum

	subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list.flist(i).fsubtotalprice
	totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list.flist(i).ftotalbuysum
	spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list.flist(i).fspendScoupon
	spendMileage_totalsum = spendMileage_totalsum + Omaechul_list.flist(i).fspendMileage
	discountEtc_totalsum = discountEtc_totalsum + Omaechul_list.flist(i).fdiscountEtc
	sumpaymentEtc_totalsum = sumpaymentEtc_totalsum + Omaechul_list.flist(i).fsumpaymentetc
	tendeliverBuysum_totalsum = tendeliverBuysum_totalsum + Omaechul_list.flist(i).ftendeliverBuysum
	tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list.flist(i).ftendeliverCount
	TTLupchepartDeliverBuySum = TTLupchepartDeliverBuySum + Omaechul_list.flist(i).fupchepartDeliverBuySum
	sunsuik_totalsum = sunsuik_totalsum + Omaechul_list.flist(i).fsunsuik
	%>
<% next %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td align="center" rowspan="2">
			�Ѱ�
		</td>
		<td align="center"  rowspan="2"><%= totalcount_totalsum %></td>
<% if (NOT C_InspectorUser) THEN %>
		<td align="right"><%= NullOrCurrFormat(TTLtotalorgitemcostsum) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalOrgDlvPay) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalorgitemcostsum-TTLtotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalCouponNotAppliedDlvPay) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalitemcostcouponNotApplied-TTLtotalitemcostsum) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalCouponNotAppliedDlvPay-TTLtotalDlvPay) %></td>

		<% IF IsNULL(TTLtotalitemcostsum) then %>
			<td align="right" colspan="2" rowspan="2"><%= CurrFormat(totalsum_totalsum) %></td>
		<% else %>
			<td align="right"><%= NullOrCurrFormat(TTLtotalitemcostsum) %></td>
			<td align="right"><%= NullOrCurrFormat(TTLtotalDlvPay) %></td>
		<% end if %>

		<% IF IsNULL(TTLtotalreducedDlvPay) then %>
			<td align="right" colspan="2" rowspan="2"><%= CurrFormat(spendScoupon_totalsum) %></td>
		<% else %>
			<td align="right"><%= NullOrCurrFormat(spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay)) %></td>
			<td align="right"><%= NullOrCurrFormat(TTLtotalDlvPay-TTLtotalreducedDlvPay) %></td>
		<% end if %>

		<td align="right" rowspan="2"><%= CurrFormat(discountEtc_totalsum) %></td>
<% end if %>
		<td align="right" rowspan="2"><%= CurrFormat((TTLtotalitemcostsum+TTLtotalDlvPay)-((spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay))+(TTLtotalDlvPay-TTLtotalreducedDlvPay))-(discountEtc_totalsum)) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(spendMileage_totalsum) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(sumpaymentEtc_totalsum) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(subtotalprice_totalsum) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(totalbuysum_totalsum) %></td>
		<td align="right"><%= CurrFormat(tendeliverBuysum_totalsum) %></td>
		<td align="center"><%= NullOrCurrFormat(TTLupchepartDeliverBuySum) %></td>
		<!-- <td align="right"><%= CurrFormat(sunsuik_totalsum) %></td>-->
		<td align="center" rowspan="2">
		<% if (subtotalprice_totalsum<>0) then %>
		    <% if Not IsNULL(sunsuik_totalsum) then %>
			<% magin_totalsum = CLNG((sunsuik_totalsum / subtotalprice_totalsum)*100*100)/100 %>
			<%= round(magin_totalsum,2) %>%
			<% end if %>
		<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<% if (NOT C_InspectorUser) THEN %>
	    <td colspan="2"><%= NullOrCurrFormat(TTLtotalorgitemcostsum+TTLtotalOrgDlvPay) %></td>
	    <td colspan="2"><%= NullOrCurrFormat((TTLtotalorgitemcostsum-TTLtotalitemcostcouponNotApplied)+(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay)) %></td>
	    <td colspan="2"><%= NullOrCurrFormat(TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay) %></td>
	    <td colspan="2"><%= NullOrCurrFormat((TTLtotalitemcostcouponNotApplied-TTLtotalitemcostsum)+(TTLtotalCouponNotAppliedDlvPay-TTLtotalDlvPay)) %></td>

	    <% IF IsNULL(TTLtotalitemcostsum) then %>
	    <% else %>
	    	<td colspan="2"><%= NullOrCurrFormat(TTLtotalitemcostsum+TTLtotalDlvPay) %></td>
	    <% end if %>

	    <% IF IsNULL(TTLtotalreducedDlvPay) then %>
	    <% else %>
	   	 <td colspan="2"><%= NullOrCurrFormat((spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay))+(TTLtotalDlvPay-TTLtotalreducedDlvPay)) %></td>
	    <% end if %>
     <% end if %>
	    <td colspan="2"><%= NullOrCurrFormat(tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" rowspan="2">
		������
		</td>
		<td align="center" rowspan="2"></td>
	<% if (NOT C_InspectorUser) THEN %>
		<td align="right" colspan="2" rowspan="2">�Һ񰡴��=&gt</td>
		<td align="center">
			<% if TTLtotalorgitemcostsum<>0 then %>
			    <%= CLNG((TTLtotalorgitemcostsum-TTLtotalitemcostcouponNotApplied)/TTLtotalorgitemcostsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if TTLtotalOrgDlvPay<>0 then %>
			    <%= CLNG((TTLtotalOrgDlvPay-(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay))/TTLtotalOrgDlvPay*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="right" colspan="2" rowspan="2">�ǸŰ����=&gt</td>
		<td align="center">
			<% if TTLtotalitemcostcouponNotApplied<>0 then %>
			    <%= CLNG((TTLtotalitemcostcouponNotApplied-TTLtotalitemcostsum)/TTLtotalitemcostcouponNotApplied*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if TTLtotalCouponNotAppliedDlvPay<>0 then %>
			    <%= CLNG((TTLtotalCouponNotAppliedDlvPay-TTLtotalDlvPay)/TTLtotalCouponNotAppliedDlvPay*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="right" colspan="2" rowspan="2">�����Ѿ״��=&gt</td>

		<% IF IsNULL(TTLtotalreducedDlvPay) then %>
		    <td align="center" colspan="2" rowspan="2">
		    <% if (totalsum_totalsum<>0) then %>
		        <%= CLNG(spendScoupon_totalsum/totalsum_totalsum*100*100)/100 %> %
		    <% end if %>
		    </td>
		<% else %>
    		<td align="center">
    		<% if TTLtotalitemcostsum<>0 then %>
    		    <%= CLNG((spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay))/TTLtotalitemcostsum*100*100)/100 %> %
    		<% end if %>
    		</td>
    		<td align="center">
    		<% if TTLtotalDlvPay<>0 then %>
    		    <%= CLNG((TTLtotalDlvPay-TTLtotalreducedDlvPay)/TTLtotalDlvPay*100*100)/100 %> %
    		<% end if %>
    		</td>
		<% end if %>

		<td align="center" rowspan="2">
			<% if totalsum_totalsum<>0 then %>
			    <%= CLNG(discountEtc_totalsum/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
	<% end if %>
		<td align="center" rowspan="2">
		</td>
		<td align="center" rowspan="2">
			<% if totalsum_totalsum<>0 then %>
			    <%= CLNG((spendMileage_totalsum)/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2">
			<% if sumpaymentEtc_totalsum<>0 then %>
			    <%= CLNG(sumpaymentEtc_totalsum/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2">
		    <% if totalsum_totalsum<>0 then %>
			    <%= CLNG(subtotalprice_totalsum/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2">
			�����Ѿ״��=&gt<br>
			<% if (totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)<>0 then %>
			    <%= CLNG(totalbuysum_totalsum/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if (totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)<>0 then %>
			    <%= CLNG(tendeliverBuysum_totalsum/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if (totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)<>0 then %>
			    <%= CLNG(TTLupchepartDeliverBuySum/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2" > </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
	<% if (NOT C_InspectorUser) THEN %>
	    <td colspan="2">
	    <% if (TTLtotalorgitemcostsum+TTLtotalOrgDlvPay)<>0 then %>
	        <%= CLNG(((TTLtotalorgitemcostsum+TTLtotalOrgDlvPay)-(TTLtotalitemcostcouponNotApplied+(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay)))/(TTLtotalorgitemcostsum+TTLtotalOrgDlvPay)*100*100)/100 %> %
	    <% end if %>
	    </td>
	    <td colspan="2">
	    <% if (TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay)<>0 then %>
	        <%= CLNG(((TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay)-(TTLtotalitemcostsum+TTLtotalDlvPay))/(TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay)*100*100)/100 %> %
	    <% end if %>
	    </td>

	    <% IF IsNULL(TTLtotalreducedDlvPay) then %>
	    <% else %>
		    <td colspan="2">
		        <% if (TTLtotalitemcostsum+TTLtotalDlvPay)<>0 then %>
		            <%= CLNG(((spendScoupon_totalsum+TTLtotalDlvPay)-((TTLtotalDlvPay-TTLtotalreducedDlvPay)+TTLtotalreducedDlvPay))/(TTLtotalitemcostsum+TTLtotalDlvPay)*100*100)/100 %> %
		        <% end if%>
		    </td>
	    <% end if %>
    <% end if %>
		<td colspan="2">
	        <% if ((totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum))<>0 then %>
	            <%= CLNG((tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
	        <% end if%>
	    </td>

	</tr>
</table>


<!-- Not Using .. OLD Ver -->
<% IF (FALSE) THEN %>
<p>------------------------<p>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if Omaechul_list.ftotalcount > 0 then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<% if dateview1 = "yes" then %>
			<td align="center" width="80">�ֹ���</td>
		<% elseif dateview1 = "no" then %>
			<td align="center" width="80">�Ա���</td>
		<% end if %>
		<% if datecancle <> "" then %>
			<td align="center" width="80">�����</td>
		<% end if %>
		<td align="center" width="70">���ֹ�<br>�Ǽ�</td>
		<td align="center" width="90">�Һ��ڰ�<br>�����Ѿ�<!-- <br>(��ۺ�����) --></td>
		<td align="center" width="80">���αݾ�</td>
		<td align="center" width="90">�ǸŰ�<br>�����Ѿ�<br>(���ΰ�)</td>
		<td align="center" width="80">��ǰ����<br>����</td>
		<td align="center"><strong>��ǰ<br>�����Ѿ�</strong></td>
		<td align="center" width="70">���ʽ�<br>����<br>����</td>
		<td align="center" width="70">���ϸ���<br>����</td>
		<td align="center" width="70">��Ÿ����</td>
		<td align="center" width="70">��ۺ�<br>����</td>
		<td align="center"><strong>�����Ѿ�<br>(�ǰ�����)<br>��ۺ�����</strong></td>
		<td align="center">���԰�<br>�Ѿ�<br>(��ǰ)</td>
		<td align="center"  width="60">�ٹ�����<br>��ۺ�<br>(����)</td>
		<td align="center"  width="60">��ü����<br>��ۺ�<br>(����)</td>
		<td align="center"><strong>�������</strong></td>
		<td align="center">����</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td></td>
        <td></td>
        <td>A</td>
        <td>B</td>
        <td>C=A-B</td>
        <td>D</td>
        <td>E=C-D</td>
        <td>F</td>
        <td>G</td>
        <td>H</td>
        <td>I</td>
        <td>J=E-(F+G+H)+I</td>
        <td>K</td>
        <td>L</td>
        <td>M</td>
        <td>N=J-(K+L+M)</td>
        <td>N/J</td>
    </tr>
	<% for i = 0 to Omaechul_list.ftotalcount -1 %>
    <tr align="center" bgcolor="#FFFFFF">
		<td align="center">
		<% if right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "�����" then %>
			<font color="blue"><%= Omaechul_list.flist(i).forderdate %></font>
		<% elseif right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "�Ͽ���" then %>
			<font color="red"><%= Omaechul_list.flist(i).forderdate %></font>
		<% else %>
			<%= Omaechul_list.flist(i).forderdate %>
		<% end if %>
		</td>
		<% if datecancle <> "" then %>
			<td align="center"><%= Omaechul_list.flist(i).fcanceldate %></td>
		<% end if %>
		<td align="center"><%= Omaechul_list.flist(i).ftotalcount %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum-Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied-Omaechul_list.flist(i).ftotalitemcostsum) %></td>

		<td align="right">
		<% if IsNULL(Omaechul_list.flist(i).ftotalitemcostsum) then %>
		<%= CurrFormat(Omaechul_list.flist(i).ftotalsum-(Omaechul_list.flist(i).ftendeliversum+Omaechul_list.flist(i).fupchepartDeliverSum)) %>
		<% else %>
		<%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostsum) %>
		<% end if %>
		</td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendScoupon) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendMileage) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdiscountEtc) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftendeliversum+Omaechul_list.flist(i).fupchepartDeliverSum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsubtotalprice) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalbuysum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftendeliverBuysum) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fupchepartDeliverBuySum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsunsuik) %></td>
		<td align="center"><%= Omaechul_list.flist(i).fmagin %>%</td>
    </tr>
	<% totalsum_totalsum = totalsum_totalsum + Omaechul_list.flist(i).ftotalsum %>
	<% totalcount_totalsum = totalcount_totalsum + Omaechul_list.flist(i).ftotalcount %>
	<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list.flist(i).fsubtotalprice %>
	<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list.flist(i).ftotalbuysum %>
	<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list.flist(i).fspendScoupon %>
	<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list.flist(i).fspendMileage %>
	<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list.flist(i).fdiscountEtc %>
	<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list.flist(i).ftendeliversum %>
	<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list.flist(i).ftendeliverCount %>
	<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list.flist(i).fsunsuik %>
	<% next %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
		�� �հ�
		</td>
		<td align="center"><%= totalcount_totalsum %></td>
		<td align="right"></td>
		<td align="right"></td>
		<td align="right"></td>
		<td align="right"></td>
		<td align="right"><%= CurrFormat(totalsum_totalsum) %></td>
		<td align="right"><%= CurrFormat(spendScoupon_totalsum) %></td>
		<td align="right"><%= CurrFormat(spendMileage_totalsum) %></td>
		<td align="right"><%= CurrFormat(discountEtc_totalsum) %></td>
		<td align="center"></td>
		<td align="right"><%= CurrFormat(subtotalprice_totalsum) %></td>
		<td align="right"><%= CurrFormat(totalbuysum_totalsum) %></td>
		<td align="right"><%= CurrFormat(tendeliversum_totalsum) %></td>
		<td align="center"></td>
		<td align="right"><%= CurrFormat(sunsuik_totalsum) %></td>
		<td align="center">
			<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>
			<%= round(magin_totalsum,2) %>%
		</td>
		<%
		totalsum_totalsum = 0
		totalcount_totalsum = 0
		subtotalprice_totalsum = 0
		totalbuysum_totalsum = 0
		spendScoupon_totalsum = 0
		spendMileage_totalsum = 0
		discountEtc_totalsum = 0
		tendeliversum_totalsum = 0
		tendeliverCount_totalsum = 0
		sunsuik_totalsum = 0
		magin_totalsum = 0
		%>
	</tr>
	<!--
	<tr bgcolor="#DDDDFF">
		<td colspan="15">
			&nbsp;&nbsp;&nbsp;���⵵ �� ���� ǥ��
			<input type=checkbox name="olddata" value="no" onclick=
			"submit();"<% if olddata="no" then  response.write "checked" %>>
		</td>
	</tr>
	//-->
	<% if (FALSE) and (olddata = "no") then %>
		<% if Omaechul_list_old.ftotalcount > 0 then %>
			<% for i = 0 to Omaechul_list_old.ftotalcount -1 %>
			<tr bgcolor="#FFFFFF">
				<td align="right">
				<% if right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "�����" then %>
					<font color="blue"><%= Omaechul_list_old.flist(i).forderdate %></font>
				<% elseif right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "�Ͽ���" then %>
					<font color="red"><%= Omaechul_list_old.flist(i).forderdate %></font>
				<% else %>
					<%= Omaechul_list_old.flist(i).forderdate %>
				<% end if %>
				</td>
				<% if datecancle <> "" then %>
					<td align="center"><%= Omaechul_list_old.flist(i).fcanceldate %></td>
				<% end if %>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalsum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftotalcount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsubtotalprice) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalbuysum) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendScoupon) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendMileage) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fdiscountEtc) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftendeliversum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftendeliverCount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsunsuik) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).fmagin %>%</td>
			</tr>
			<% totalsum_totalsum = totalsum_totalsum + Omaechul_list_old.flist(i).ftotalsum %>
			<% totalcount_totalsum = totalcount_totalsum + Omaechul_list_old.flist(i).ftotalcount %>
			<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list_old.flist(i).fsubtotalprice %>
			<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list_old.flist(i).ftotalbuysum %>
			<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list_old.flist(i).fspendScoupon %>
			<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list_old.flist(i).fspendMileage %>
			<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list_old.flist(i).fdiscountEtc %>
			<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list_old.flist(i).ftendeliversum %>
			<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list_old.flist(i).ftendeliverCount %>
			<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list_old.flist(i).fsunsuik %>
			<% next %>
			<tr bgcolor="#F4F4F4">
				<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
				�� �հ�
				</td>
				<td align="right">
					<%= CurrFormat(totalsum_totalsum) %>
				</td>
				<td align="center">
					<%= totalcount_totalsum %>
				</td>
				<td align="right">
					<%= CurrFormat(subtotalprice_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(totalbuysum_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(spendScoupon_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(spendMileage_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(discountEtc_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(tendeliversum_totalsum) %>
				</td>
				<td align="center">
					<%= CurrFormat(tendeliverCount_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(sunsuik_totalsum) %>
				</td>
				<td align="center">
					<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>
					<%= round(magin_totalsum,2) %>%
				</td>
			</tr>
		<% else %>
			<tr align="center" bgcolor="#DDDDFF">
		    	<td align=center bgcolor="#FFFFFF" colspan="15">���⵵ �˻� ����� �����ϴ�.</td>
		    </tr>
		<% end if %>
	<% end if %>

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>


</table>
<% end if %>

<%
	set Omaechul_list = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
