<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ü��� ��չ���� ����
' History : 2007.08.03 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->

<%
dim yyyy , mm , checkmode2 , checkmode1 , checkmode3,disp , disexcel , i
	yyyy = request("yyyy1")
	mm = request("mm1")
	checkmode1 = request("checkmode1")
	checkmode2 = request("checkmode2")
	checkmode3 = request("checkmode3")
	disp = request("disp")

	if disp="" then disp="A"

if disp = "A" then
dim onomalitemsummary
	set onomalitemsummary = new Cchulgoitemlist
	onomalitemsummary.frectyyyy = yyyy
	onomalitemsummary.frectmm = mm
	onomalitemsummary.fnomalitemsummary()

dim ojumunitemsummary
	set ojumunitemsummary = new Cchulgoitemlist
	ojumunitemsummary.frectyyyy = yyyy
	ojumunitemsummary.frectmm = mm
	ojumunitemsummary.fjumunitemsummary()

dim onomalmakeridsummary
	set onomalmakeridsummary = new Cchulgoitemlist
	onomalmakeridsummary.frectyyyy = yyyy
	onomalmakeridsummary.frectmm = mm
	onomalmakeridsummary.fnomalmakeridsummary()

dim ojumunmakeridsummary
	set ojumunmakeridsummary = new Cchulgoitemlist
	ojumunmakeridsummary.frectyyyy = yyyy
	ojumunmakeridsummary.frectmm = mm
	ojumunmakeridsummary.fjumunmakeridsummary()
end if

if disp = "B" then
	dim omidalitem
		set omidalitem = new Cchulgoitemlist
		omidalitem.frectyyyy = yyyy
		omidalitem.frectmm = mm
		omidalitem.fupcheitemmidal()
end if

if disp = "C" then
dim omidalmaker
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
function ExcelSheet(yyyy,mm,checkmode1,checkmode2,checkmode3){
	var excel = window.open('/admin/chulgo/upchebaesonglist_excel.asp?yyyy='+yyyy+'&mm='+mm+'&checkmode1='+checkmode1+'&checkmode2='+checkmode2+'&checkmode3='+checkmode3,'excelsheet','width=1024,height=768,scrollbars=yes,resizable=yes');
	excel.focus();
}

</script>

<!-- �������Ϸ� ���� ��� �κ� -->
<%
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"upchechulgo_"+yyyy+"_"+mm+".xls"
%>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr width="50">
	<td>
		<font color="red"><strong>��ü �����Ȳ</strong></font>
	</td>
</tr>
</table>
<!--ǥ ��峡-->

<!-- ǥ �˻��κ� ����-->


<% dim totald0,totald1,totald2,totald3,totald4,totald5,totald6,totald7,totald8,totald9,totald10,totald11,totald12 , totalcount,totaldiv4_1,totaldiv4_2 %>
	<% if disp = "A" then %>
		<% if onomalitemsummary.flist(i).fitemd0 <> "" then %>

		<!-- ��ǰ�� ��� �ҿ���(�Ϲݻ�ǰ) ����-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				��ǰ�� ��� �ҿ���[�Ϲݻ�ǰ]
			</td>
			<td bgcolor="ffffff" colspan=9>
				��ǥ : ���ع̴� ��ǰ 5% �̳�
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">����</td>
			<td colspan=4><div align="center">��������</td>
			<td colspan=4><div align="center">���ع̴�</td>
			<td rowspan=2><div align="center">�հ�</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+0</td>
					<td><div align="center">D+1</td>
					<td><div align="center">D+2</td>
					<td><div align="center">D+3</td>
					<td><div align="center">D+4</td>
					<td><div align="center">D+5</td>
					<td><div align="center">D+6</td>
					<td><div align="center">D+7�̻�</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">���Ǽ�</td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(onomalitemsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3+onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">����</td>
			<td><div align="center"><% totald0 = (onomalitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (onomalitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (onomalitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (onomalitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (onomalitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (onomalitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (onomalitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (onomalitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2>100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><div align="center"><% totaldiv4_1 = ((onomalitemsummary.flist(i).fitemd0+onomalitemsummary.flist(i).fitemd1+onomalitemsummary.flist(i).fitemd2+onomalitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><div align="center"><% totaldiv4_2 = ((onomalitemsummary.flist(i).fitemd4+onomalitemsummary.flist(i).fitemd5+onomalitemsummary.flist(i).fitemd6+onomalitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- ��ǰ�� ��� �ҿ���(�Ϲݻ�ǰ) ��-->

		<!-- ��ǰ�� ��� �ҿ���(�ֹ����ۻ�ǰ) ����-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				��ǰ�� ��� �ҿ���[�ֹ����ۻ�ǰ]
			</td>
			<td bgcolor="ffffff" colspan=9>
				��ǥ : ���ع̴� ��ǰ 5% �̳�
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">����</td>
			<td colspan=4><div align="center">��������</td>
			<td colspan=4><div align="center">���ع̴�</td>
			<td rowspan=2><div align="center">�հ�</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+5����</td>
					<td><div align="center">D+6</td>
					<td><div align="center">D+7</td>
					<td><div align="center">D+8</td>
					<td><div align="center">D+9</td>
					<td><div align="center">D+10</td>
					<td><div align="center">D+11</td>
					<td><div align="center">D+12</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">���Ǽ�</td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(ojumunitemsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3+ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">����</td>
			<td><div align="center"><% totald0 = (ojumunitemsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (ojumunitemsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (ojumunitemsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (ojumunitemsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (ojumunitemsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (ojumunitemsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (ojumunitemsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (ojumunitemsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2>100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><div align="center"><% totaldiv4_1 = ((ojumunitemsummary.flist(i).fitemd0+ojumunitemsummary.flist(i).fitemd1+ojumunitemsummary.flist(i).fitemd2+ojumunitemsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><div align="center"><% totaldiv4_2 = ((ojumunitemsummary.flist(i).fitemd4+ojumunitemsummary.flist(i).fitemd5+ojumunitemsummary.flist(i).fitemd6+ojumunitemsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- ��ǰ�� ��� �ҿ���(�ֹ����ۻ�ǰ) ��-->

		<!-- �귣�庰 ��� ��� �ҿ��� (�Ϲݻ�ǰ) ����-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				�귣�庰 ��� ��� �ҿ��� [�Ϲݻ�ǰ]
			</td>
			<td bgcolor="ffffff" colspan=9>
				��ǥ : ���ع̴� �귣�� 5% �̳�
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">����</td>
			<td colspan=4><div align="center">��������</td>
			<td colspan=4><div align="center">���ع̴�</td>
			<td rowspan=2><div align="center">�հ�</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+0</td>
					<td><div align="center">D+1����</td>
					<td><div align="center">D+2����</td>
					<td><div align="center">D+3����</td>
					<td><div align="center">D+3�ʰ�</td>
					<td><div align="center">D+4�ʰ�</td>
					<td><div align="center">D+5�ʰ�</td>
					<td><div align="center">D+6�ʰ�</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">���Ǽ�</td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(onomalmakeridsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3+onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">����</td>
			<td><div align="center"><% totald0 = (onomalmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (onomalmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (onomalmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (onomalmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (onomalmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (onomalmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (onomalmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (onomalmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2><div align="center">100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><% totaldiv4_1 = ((onomalmakeridsummary.flist(i).fitemd0+onomalmakeridsummary.flist(i).fitemd1+onomalmakeridsummary.flist(i).fitemd2+onomalmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><% totaldiv4_2 = ((onomalmakeridsummary.flist(i).fitemd4+onomalmakeridsummary.flist(i).fitemd5+onomalmakeridsummary.flist(i).fitemd6+onomalmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- �귣�庰 ��� ��� �ҿ��� (�Ϲݻ�ǰ) ��-->

		<!-- �귣�庰 ��� ��� �ҿ��� (���ۻ�ǰ) ����-->
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4">
				�귣�庰 ��� ��� �ҿ��� [���ۻ�ǰ]
			</td>
			<td bgcolor="ffffff" colspan=9>
				��ǥ : ���ع̴� �귣�� 5% �̳�
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td rowspan=2><div align="center">����</td>
			<td colspan=4><div align="center">��������</td>
			<td colspan=4><div align="center">���ع̴�</td>
			<td rowspan=2><div align="center">�հ�</td>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">D+5����</td>
					<td><div align="center">D+6����</td>
					<td><div align="center">D+7����</td>
					<td><div align="center">D+8����</td>
					<td><div align="center">D+8�ʰ�</td>
					<td><div align="center">D+9�ʰ�</td>
					<td><div align="center">D+10�ʰ�</td>
					<td><div align="center">D+11�ʰ�</td>
				</tr>
		</tr>
		<tr bgcolor="ffffff">
			<td><div align="center">���Ǽ�</td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd0) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd1) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd2) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd3) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd4) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd5) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd6) %></td>
			<td><div align="center"><%= CurrFormat(ojumunmakeridsummary.flist(i).fitemd7) %></td>
			<td><div align="center"><% totalcount = ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3+ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7 %>
			<%= CurrFormat(totalcount) %></td>
		</tr>
		<tr bgcolor="ffffff">
			<td rowspan=2><div align="center">����</td>
			<td><div align="center"><% totald0 = (ojumunmakeridsummary.flist(i).fitemd0 / totalcount)*100 %><%= round(totald0,2) %>%</td>
			<td><div align="center"><% totald1 = (ojumunmakeridsummary.flist(i).fitemd1 / totalcount)*100 %><%= round(totald1,2) %>%</td>
			<td><div align="center"><% totald2 = (ojumunmakeridsummary.flist(i).fitemd2 / totalcount)*100 %><%= round(totald2,2) %>%</td>
			<td><div align="center"><% totald3 = (ojumunmakeridsummary.flist(i).fitemd3 / totalcount)*100 %><%= round(totald3,2) %>%</td>
			<td><div align="center"><% totald4 = (ojumunmakeridsummary.flist(i).fitemd4 / totalcount)*100 %><%= round(totald4,2) %>%</td>
			<td><div align="center"><% totald5 = (ojumunmakeridsummary.flist(i).fitemd5 / totalcount)*100 %><%= round(totald5,2) %>%</td>
			<td><div align="center"><% totald6 = (ojumunmakeridsummary.flist(i).fitemd6 / totalcount)*100 %><%= round(totald6,2) %>%</td>
			<td><div align="center"><% totald7 = (ojumunmakeridsummary.flist(i).fitemd7 / totalcount)*100 %><%= round(totald7,2) %>%</td>
			<td rowspan=2>100%</td>
				<tr bgcolor="ffffff">
					<td colspan=4><div align="center"><% totaldiv4_1 = ((ojumunmakeridsummary.flist(i).fitemd0+ojumunmakeridsummary.flist(i).fitemd1+ojumunmakeridsummary.flist(i).fitemd2+ojumunmakeridsummary.flist(i).fitemd3)/totalcount)*100 %><%= round(totaldiv4_1,2) %>%</td>
					<td colspan=4><div align="center"><% totaldiv4_2 = ((ojumunmakeridsummary.flist(i).fitemd4+ojumunmakeridsummary.flist(i).fitemd5+ojumunmakeridsummary.flist(i).fitemd6+ojumunmakeridsummary.flist(i).fitemd7)/totalcount)*100 %><%= round(totaldiv4_2,2) %>%</td>
				</tr>
		</tr>
		</table>
		<!-- �귣�庰 ��� ��� �ҿ��� (���ۻ�ǰ) ��-->
	<% else %>
		<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	    	<td align=center bgcolor="#FFFFFF">�˻� ����� �����ϴ�.</td>
	    </tr>
		</table>
<% end if %>

	<% end if %>

	<!-- ���ع̴� ��ǰ ���� -->
	<% if disp = "B" then %>
		<% if omidalitem.FTotalCount > 1 then %>
			<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" align="center">
			<tr>
				<td  bgcolor="F4F4F4" colspan=2>
				���� �̴� ��ǰ
				</td>
				<td bgcolor="ffffff" colspan=9>
				</td>
			</tr>
			<tr bgcolor=#DDDDFF>
				<td><div align="center">�귣��id</td>
				<td><div align="center">��ǰ�ڵ�</td>
				<td><div align="center">��ǰ��</td>
				<td><div align="center">��ǰ����</td>
				<td><div align="center">��չ����</td>
				<td><div align="center">��۰Ǽ�</td>
			</tr>

			<% for i=0 to omidalitem.FTotalCount - 1 %>
				<% ''if omidalitem.flist(i).favgdlvdate > 3 and omidalitem.flist(i).fdelivercount >=10 then %>
					<tr bgcolor="ffffff">
						<td><div align="center"><%= omidalitem.flist(i).fmakerid %></td>
						<td><div align="center"><%= omidalitem.flist(i).fitemid %></td>
						<td><div align="center"><%= omidalitem.flist(i).fitemname %></td>
						<td><div align="center"><%= omidalitem.flist(i).fitemdivname %></td>
						<td><div align="center"><%= round(omidalitem.flist(i).favgdlvdate,2) %></td>
						<td><div align="center"><%= omidalitem.flist(i).fdelivercount %></td>
					</tr>
				<% ''end if %>
			<% next %>
			</table>
			<% else %>
			<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr align="center" bgcolor="#DDDDFF">
				<td align=center bgcolor="#FFFFFF">�˻� ����� �����ϴ�.</td>
			</tr>
			</table>
	<% end if %>
<% end if %>
	<!-- ���ع̴� ��ǰ ��-->

	<!-- ���ع̴� �귣��,�Ϲݻ�ǰ,�ֹ����ۻ�ǰ����-->
	<% if disp = "C" then %>
		<% if omidalmaker.ftotalcount > 0 then %>
		<table border="0" class="a" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td width="50%">

				<!--���ع̴޺귣��,�Ϲݻ�ǰ����-->
				<table border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" width="100%">
				<tr>
					<td  bgcolor="F4F4F4" colspan=2>
					���� �̴޺귣��[�Ϲݻ�ǰ]
					</td>
					<td bgcolor="ffffff" colspan=2>
					</td>
				</tr>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">���</td>
					<td><div align="center">�귣��id</td>
					<td><div align="center">��չ����</td>
					<td><div align="center">��۰Ǽ�</td>
				</tr>

				<% for i=0 to omidalmaker.FTotalCount - 1 %>
					<% if omidalmaker.flist(i).fitemdiv = "01" then %>
						<tr bgcolor="ffffff">
							<td><div align="center"><%= omidalmaker.flist(i).fyyyy %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fmakerid %></td>
							<td><div align="center"><%= round(omidalmaker.flist(i).favgdlvdate,2) %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fdelivercount %></td>
						</tr>
					<% end if %>
				<% next %>
				</table>
				<!--���ع̴޺귣��,�Ϲݻ�ǰ��-->
			</td>
			<td width="50%">
				<!--���ع̴޺귣��,�ֹ����ۻ�ǰ����-->
				<table border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" width="100%">
				<tr>
					<td  bgcolor="F4F4F4" colspan=2>
					���� �̴޺귣��[�ֹ����ۻ�ǰ]
					</td>
					<td bgcolor="ffffff" colspan=2>

					</td>
				</tr>
				<tr bgcolor=#DDDDFF>
					<td><div align="center">���</td>
					<td><div align="center">�귣��id</td>
					<td><div align="center">��չ����</td>
					<td><div align="center">��۰Ǽ�</td>
				</tr>

				<% for i=0 to omidalmaker.FTotalCount - 1 %>
					<% if omidalmaker.flist(i).fitemdiv = "06" or omidalmaker.flist(i).fitemdiv = "16"  then %>
						<tr bgcolor="ffffff">
							<td><div align="center"><%= omidalmaker.flist(i).fyyyy %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fmakerid %></td>
							<td><div align="center"><%= round(omidalmaker.flist(i).favgdlvdate,2) %></td>
							<td><div align="center"><%= omidalmaker.flist(i).fdelivercount %></td>
						</tr>
					<% end if %>
				<% next %>
				</table>
				<!--���ع̴޺귣��,�ֹ����ۻ�ǰ��-->
			</td>
		</tr>
	</table>
	<% else %>
		<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	    <tr align="center" bgcolor="#DDDDFF">
	    	<td align=center bgcolor="#FFFFFF">�˻� ����� �����ϴ�.</td>
	    </tr>
		</table>
	<% end if %>
<% end if %>
<!-- ���ع̴� �귣��,�Ϲݻ�ǰ,�ֹ����ۻ�ǰ��-->

<%
set omidalitem = nothing
set omidalmaker = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->

