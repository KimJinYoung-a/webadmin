<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ü��� ��չ���� ����
' History : 2007.08.03 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->
<%
dim yyyy , mm ,disp , disexcel , i ,omidalmaker
dim onomalitemsummary ,ojumunitemsummary ,onomalmakeridsummary ,ojumunmakeridsummary ,omidalitem,totaldiv4_1,totaldiv4_2
dim totald0,totald1,totald2,totald3,totald4,totald5,totald6,totald7,totald8,totald9,totald10,totald11,totald12 , totalcount
dim nn, dsum, makerid
	yyyy = request("yyyy1")
	mm = request("mm1")
	if (yyyy="") then yyyy = Cstr(Year(now()))		'�˻�â�� �⺻������ �̹��⵵�� �ִ´�
	if (mm="") then mm = Cstr(Month(now()))			'�˻�â�� �⺻������ �̹����� �ִ´�
	disp = request("disp")
	if disp="" then disp="A"						'�˻�â�� �⺻������ ���(A)�� �����Ѵ�.
	menupos = request("menupos")
    makerid = requestCheckvar(request("makerid"),32)

'��輱�ý�
if disp = "A" then
	set onomalitemsummary = new Cchulgoitemlist
		onomalitemsummary.frectyyyy = yyyy
		onomalitemsummary.frectmm = mm
		onomalitemsummary.fnomalitemsummary()

	set ojumunitemsummary = new Cchulgoitemlist
		ojumunitemsummary.frectyyyy = yyyy
		ojumunitemsummary.frectmm = mm
		ojumunitemsummary.fjumunitemsummary()

	set onomalmakeridsummary = new Cchulgoitemlist
		onomalmakeridsummary.frectyyyy = yyyy
		onomalmakeridsummary.frectmm = mm
		onomalmakeridsummary.fnomalmakeridsummary()

	set ojumunmakeridsummary = new Cchulgoitemlist
		ojumunmakeridsummary.frectyyyy = yyyy
		ojumunmakeridsummary.frectmm = mm
		ojumunmakeridsummary.fjumunmakeridsummary()
end if

'���ع̴޻�ǰ ���ý�
if disp = "B" then
	set omidalitem = new Cchulgoitemlist
		omidalitem.frectyyyy = yyyy
		omidalitem.frectmm = mm
		omidalitem.FrectMakerid=makerid
		omidalitem.fupcheitemmidal()
end if

'���ع̴޺귣�� ���ý�
if disp = "C" then
	set omidalmaker = new Cchulgoitemlist
		omidalmaker.frectyyyy = yyyy
		omidalmaker.frectmm = mm
		omidalmaker.fupcheitemmidalmaker()
end if
%>

<script language="javascript">

function formsubmit(frm){
	frm.submit();
}

//������� ����
function ExcelSheet(yyyy,mm,disp){
	var excel = window.open('/admin/chulgo/upchebaesonglist_excel.asp?yyyy1='+yyyy+'&mm1='+mm+'&disp='+disp,'excelsheet','width=1024,height=768,scrollbars=yes,resizable=yes');
	excel.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
   		��: &nbsp;<% DrawYMBox yyyy,mm %>
    	<input type="radio" name="disp" value="A" <% if disp="A" then response.write "checked" %>>���
    	<input type="radio" name="disp" value="B" <% if disp="B" then response.write "checked" %>>���ع̴� ��ǰ
    	<input type="radio" name="disp" value="C" <% if disp="C" then response.write "checked" %>>���ع̴� �귣��
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	<% if disp="B" then %>
	<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
	<% end if %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	    * ���޳��� ���� �˻� �����մϴ�.<br>
		* ��������Ϻ��� ����ϱ��� �Ⱓ�� ����(D+0 ������� : D+1 1�������),<b>������ ����</b>, �Ϲݻ�ǰ : ��������Ϻ��� 3�ϳ� ���, �ֹ����ۻ�ǰ :��������Ϻ��� 8���̳� ���)
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<% if disp = "A" then %>
	<% if onomalitemsummary.flist(i).fitemd0 <> 0 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<!-- ��ǰ�� ��� �ҿ���(�Ϲݻ�ǰ) ����-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			��ǰ�� ��� �ҿ���[�Ϲݻ�ǰ]
		</td>
		<td colspan=9>
			��ǥ : ���ع̴� ��ǰ 5% �̳�&nbsp; &nbsp;
			<input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>','<%=disp%>')" value="������ ���" class="button">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>����</td>
		<td colspan=4>��������</td>
		<td colspan=4>���ع̴�</td>
		<td rowspan=2>�հ�</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+0</td>
		<td>D+1</td>
		<td>D+2</td>
		<td>D+3</td>
		<td>D+4</td>
		<td>D+5</td>
		<td>D+6</td>
		<td>D+7�̻�</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>���Ǽ�</td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(onomalitemsummary.flist(i).fitemd7) %></td>
		<td><% totalcount = onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3+onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7 %>
		<%= CurrFormat(totalcount) %></td>
	</tr>
	<% if (totalcount<>0) then %>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>����</td>
		<td><% totald0 = (onomalitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (onomalitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (onomalitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (onomalitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (onomalitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (onomalitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (onomalitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (onomalitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
    <% end if %>

	<!-- ��ǰ�� ��� �ҿ���(�ֹ����ۻ�ǰ) ����-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			��ǰ�� ��� �ҿ���[�ֹ�����(����)��ǰ]
		</td>
		<td bgcolor="ffffff" colspan=9>
			��ǥ : ���ع̴� ��ǰ 5% �̳�
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>����</td>
		<td colspan=4>��������</td>
		<td colspan=4>���ع̴�</td>
		<td rowspan=2>�հ�</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+5����</td>
		<td>D+6</td>
		<td>D+7</td>
		<td>D+8</td>
		<td>D+9</td>
		<td>D+10</td>
		<td>D+11</td>
		<td>D+12</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>���Ǽ�</td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(ojumunitemsummary.flist(i).fitemd7) %></td>
		<td>
			<% totalcount = ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3+ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %>
		</td>
	</tr>
	<% if (totalcount<>0) then %>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>����</td>
		<td><% totald0 = (ojumunitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (ojumunitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (ojumunitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (ojumunitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (ojumunitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (ojumunitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (ojumunitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (ojumunitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
    <% end if %>

	<!-- �귣�庰 ��� ��� �ҿ��� (�Ϲݻ�ǰ) ����-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			�귣�庰 ��� ��� �ҿ��� [�Ϲݻ�ǰ]
		</td>
		<td bgcolor="ffffff" colspan=9>
			��ǥ : ���ع̴� �귣�� 5% �̳�
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>����</td>
		<td colspan=4>��������</td>
		<td colspan=4>���ع̴�</td>
		<td rowspan=2>�հ�</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+0</td>
		<td>D+1����</td>
		<td>D+2����</td>
		<td>D+3����</td>
		<td>D+3�ʰ�</td>
		<td>D+4�ʰ�</td>
		<td>D+5�ʰ�</td>
		<td>D+6�ʰ�</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>�귣���</td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd7) %></td>
		<td>
			<% totalcount = onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3+onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %>
		</td>
	</tr>
	<% if (totalcount<>0) then %>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>����</td>
		<td><% totald0 = (onomalmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (onomalmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (onomalmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (onomalmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (onomalmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (onomalmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (onomalmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (onomalmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
    <% end if %>

	<!-- �귣�庰 ��� ��� �ҿ��� (���ۻ�ǰ) ����-->
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td>
			�귣�庰 ��� ��� �ҿ��� [�ֹ�����(����)��ǰ]
		</td>
		<td bgcolor="ffffff" colspan=9>
			��ǥ : ���ع̴� �귣�� 5% �̳�
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan=2>����</td>
		<td colspan=4>��������</td>
		<td colspan=4>���ع̴�</td>
		<td rowspan=2>�հ�</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>D+5����</td>
		<td>D+6����</td>
		<td>D+7����</td>
		<td>D+8����</td>
		<td>D+8�ʰ�</td>
		<td>D+9�ʰ�</td>
		<td>D+10�ʰ�</td>
		<td>D+11�ʰ�</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td>�귣���</td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd0) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd1) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd2) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd3) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd4) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd5) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd6) %></td>
		<td><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd7) %></td>
		<td><% totalcount = ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3+ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7 %>
		<%= CurrFormat(totalcount) %></td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td rowspan=2>����</td>
		<td><% totald0 = (ojumunmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
		<td><% totald1 = (ojumunmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
		<td><% totald2 = (ojumunmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
		<td><% totald3 = (ojumunmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
		<td><% totald4 = (ojumunmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
		<td><% totald5 = (ojumunmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
		<td><% totald6 = (ojumunmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
		<td><% totald7 = (ojumunmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr bgcolor="ffffff" align="center">
		<td colspan=4><% totaldiv4_1 = ((ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
		<td colspan=4><% totaldiv4_2 = ((ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
	</tr>
	</table>

<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�˻� ����� �����ϴ�</td>
    </tr>
</table>
<% end if %>

<%
'/���ع̴� ��ǰ ����
elseif disp = "B" then

%>
	<% if omidalitem.FTotalCount > 1 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" colspan=2>
			���� �̴� ��ǰ
		</td>
		<td bgcolor="ffffff" colspan=9>
			<!-- ��۰Ǽ� 10ȸ�̻� ��ǰ ���� -->
			&nbsp; &nbsp;
			<input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>','<%=disp%>')" value="������ ���" class="button">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�귣��id</td>
		<td>��ǰ�ڵ�</td>
		<td>��ǰ��</td>
		<td>��ǰ����</td>
		<td>��չ����</td>
		<td>��۰Ǽ�</td>
	</tr>

	<% for i=0 to omidalitem.FTotalCount - 1 %>
    <%
        nn=nn+1
		dsum=dsum+omidalitem.flist(i).fdelivercount
	%>
	<tr bgcolor="ffffff">
		<td><%= omidalitem.flist(i).fmakerid %></td>
		<td><%= omidalitem.flist(i).fitemid %></td>
		<td><%= omidalitem.flist(i).fitemname %></td>
		<td><%= omidalitem.flist(i).fitemdivname %></td>
		<td><%= CLng(omidalitem.flist(i).favgdlvdate*100)/100 %></td>
		<td><%= (omidalitem.flist(i).fdelivercount) %></td>
	</tr>

	<% next %>
	<tr bgcolor="#EEEEEE" align="center">
	    <td>�Ѱ�</td>
	    <td><%=FormatNumber(nn,0)%></td>
	    <td></td>
	    <td></td>
	    <td></td>
	    <td><%=FormatNumber(dsum,0)%></td>
	</tr>
	</table>

	<% else %>

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td>�˻� ����� �����ϴ�.</td>
	</tr>
	</table>
	<% end if %>

<%
'//���ع̴� �귣��,�Ϲݻ�ǰ,�ֹ����ۻ�ǰ����
elseif disp = "C" then
%>
	<% if omidalmaker.ftotalcount > 0 then %>
	<table border="0" class="a" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td width="49%">
			<!--���ع̴޺귣��,�Ϲݻ�ǰ����-->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td colspan=2>
					���� �̴޺귣��[�Ϲݻ�ǰ]
				</td>
				<td colspan=2>
					&nbsp; &nbsp;
					<input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>','<%=disp%>')" value="������ ���" class="button">
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>���</td>
				<td>�귣��id</td>
				<td>��չ����</td>
				<td>��۰Ǽ�</td>
			</tr>
			<%
			nn = 0
			dsum = 0
			for i=0 to omidalmaker.FTotalCount - 1
			if omidalmaker.flist(i).fitemdiv = "01" then
			    nn=nn+1
			    dsum=dsum+omidalmaker.flist(i).fdelivercount
			%>
		    <tr bgcolor="ffffff" align="center">
				<td><%= omidalmaker.flist(i).fyyyy %></td>
				<td><a href="?disp=B&makerid=<%= omidalmaker.flist(i).fmakerid %>&yyyy1=<%=yyyy%>&mm1=<%=mm%>"><%= omidalmaker.flist(i).fmakerid %></a></td>
				<td><%= CLNG(omidalmaker.flist(i).favgdlvdate*100)/100 %></td>
				<td><%= omidalmaker.flist(i).fdelivercount %></td>
			</tr>
			<%
			end if
			next
			%>
			<tr bgcolor="#EEEEEE" align="center">
			    <td>�Ѱ�</td>
			    <td><%=FormatNumber(nn,0)%></td>
			    <td></td>
			    <td><%=FormatNumber(dsum,0)%></td>
			</tr>
			</table>
		</td>
		<td width="1%"></td>
		<td width="49%" valign="top">
			<!--���ع̴޺귣��,�ֹ����ۻ�ǰ����-->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td colspan=2>
				���� �̴޺귣��[�ֹ����ۻ�ǰ]
				</td>
				<td colspan=2>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>���</td>
				<td>�귣��id</td>
				<td>��չ����</td>
				<td>��۰Ǽ�</td>
			</tr>
			<%
			nn = 0
			dsum = 0
			for i=0 to omidalmaker.FTotalCount - 1
			if omidalmaker.flist(i).fitemdiv = "06" or omidalmaker.flist(i).fitemdiv = "16" then
			    nn=nn+1
			    dsum=dsum+omidalmaker.flist(i).fdelivercount
			%>
			<tr bgcolor="ffffff" align="center">
				<td><%= omidalmaker.flist(i).fyyyy %></td>
				<td><a href="?disp=B&makerid=<%= omidalmaker.flist(i).fmakerid %>&yyyy1=<%=yyyy%>&mm1=<%=mm%>"><%= omidalmaker.flist(i).fmakerid %></a></td>
				<td><%= CLNG(omidalmaker.flist(i).favgdlvdate*100)/100 %></td>
				<td><%= omidalmaker.flist(i).fdelivercount %></td>
			</tr>
			<%
		    end if
			next
			%>
			<tr bgcolor="#EEEEEE" align="center">
			    <td>�Ѱ�</td>
			    <td><%=FormatNumber(nn,0)%></td>
			    <td></td>
			    <td><%=FormatNumber(dsum,0)%></td>
			</tr>
			</table>
		</td>
	</tr>
	</table>

	<% else %>

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	    <tr align="center" bgcolor="#FFFFFF">
	    	<td>�˻� ����� �����ϴ�.</td>
	    </tr>
	</table>
	<% end if %>
<% end if %>

<%
'��輱�ý�
if disp = "A" then
	set onomalitemsummary = nothing
	set ojumunitemsummary = nothing
	set onomalmakeridsummary = nothing
	set ojumunmakeridsummary = nothing
end if

'���ع̴޻�ǰ ���ý�
if disp = "B" then
	set omidalitem = nothing
end if

'���ع̴޺귣�� ���ý�
if disp = "C" then
	set omidalmaker = nothing
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
