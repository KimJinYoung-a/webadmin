<%@ language=vbscript %>
<% option explicit
Response.CharSet = "euc-kr"%>
<%
'###########################################################
' Description :  �Ϻ������ ����
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
dim frectbaljutotalno, frectrectbaesong,frectcentertotalno,frectcancelno,frecttotalchulgono, frectdelay0chulgo
dim frectdelay1chulgo,frectdelay2chulgo,frectdelay3over,frectrectdaychulgo,frectbaesongtotal
dim ffrectdelay0chulgo,ffrectdelay1chulgo,ffrectdelay2chulgo,ffrectdelay3chulgo ,yyyy , mm
dim frectdangilchulgo1 ,frectdangilchulgo2,frectdangilchulgo3,frectdangilchulgo4,frectdangilchulgo5,frectdangilchulgo6
	yyyy = request("yyyy1")
	mm = request("mm1")
		if (yyyy="") then yyyy = Cstr(Year(now()))
		if (mm="") then mm = Cstr(Month(now()))

dim ochulgo , i
	set ochulgo = new Cchulgoitemlist
	ochulgo.frectyyyy = yyyy
	ochulgo.frectmm = mm
	ochulgo.fchulgoitemlist()

dim ochulgomonth
	set ochulgomonth = new Cchulgoitemlist
	ochulgomonth.frectyyyy = yyyy
	ochulgomonth.frectmm = mm
	ochulgomonth.fchulgomonth()
%>

<script language="javascript">

//������� ����
function ExcelSheet(yyyy,mm){
	var excel = window.open('/admin/chulgo/chulgolist_excel.asp?yyyy1=' + yyyy + ' &mm1=' +mm ,'excelsheet','width=1024,height=768,scrollbars=yes,resizable=yes');
	excel.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��: &nbsp;<% DrawYMBox yyyy,mm %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		��������ñ��� : ����3�� �����Ϸ��(�������Աݰ�)���� �������, ��������� �������� ����ϱ��� �ҿ���ϼ� ���� , ��������� ��� ������ �ݿ��Ǿ� ���� ����<br>
		��<font color="red">�̿����</font> �Ǵ� ������ ���Ե��� ����.<br>
		��<font color="red">��ġ����Ŀ ���ⱸ����</font> �ֹ��� ���Ե��� ����.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<% if ochulgo.FTotalCount > 0 then %>
<!-- �Ϻ� �����Ȳ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="ffffff">
	<td  colspan=2>
	�Ϻ� ��� ��Ȳ
	</td>
	<td colspan=9>
		���� ����� ��ǥ : 99% <input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>')" value="������ ���" class="button">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2">
		��¥
	</td>
	<td rowspan="2">
		��������ðǼ�
	</td>
	<td rowspan="2">
		��ü��ۺ���
	</td>
	<td colspan=3>
		��ü��۰Ǽ�
	</td>
	<td colspan=4>
		�����
	</td>
	<td rowspan="2">
		���������
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>
		�ѰǼ�
	</td>
	<td>
		��ҰǼ�
	</td>
	<td>
		���Ǽ�
	</td>
	<td>
		�������
	</td>
	<td>
		1������
	</td>
	<td>
		2������
	</td>
	<td>
		3���̻�
	</td>
</tr>
<%
for i=0 to ochulgo.FTotalCount - 1
%>
<tr bgcolor="ffffff" align="center">
	<td>
		<%= ochulgo.flist(i).fmm %>�� <%= ochulgo.flist(i).fdd %>��
	</td>
	<td>
		<% frectbaljutotalno = frectbaljutotalno+ochulgo.flist(i).fbaljutotalno %>
		<%= CurrFormat(ochulgo.flist(i).fbaljutotalno) %>
	</td>
	<td>
		<%= round(ochulgo.flist(i).frectbaesong,1) %>%
		<% frectrectbaesong = frectrectbaesong+ochulgo.flist(i).frectbaesong %>
	</td>
	<td>
		<%= CurrFormat(ochulgo.flist(i).fcentertotalno)	%>
		<% frectcentertotalno = frectcentertotalno+ochulgo.flist(i).fcentertotalno %>
	</td>
	<td>
		<%= ochulgo.flist(i).fcancelno %>
		<% frectcancelno = frectcancelno+ochulgo.flist(i).fcancelno %>
	</td>
	<td>
		<%= CurrFormat(ochulgo.flist(i).ftotalchulgono) %>
		<% frecttotalchulgono = frecttotalchulgono+ochulgo.flist(i).ftotalchulgono %>
	</td>
	<td>
		<font color="red"><%= CurrFormat(ochulgo.flist(i).fdelay0chulgo) %></font><% frectdelay0chulgo = frectdelay0chulgo+ochulgo.flist(i).fdelay0chulgo %>
	</td>
	<td>
		<%= ochulgo.flist(i).fdelay1chulgo %><% frectdelay1chulgo = frectdelay1chulgo+ochulgo.flist(i).fdelay1chulgo %>
	</td>
	<td>
		<%= ochulgo.flist(i).fdelay2chulgo %><% frectdelay2chulgo = frectdelay2chulgo+ochulgo.flist(i).fdelay2chulgo %>
	</td>
	<td>
		<font color="red"><%= ochulgo.flist(i).fdelay3over %></font><% frectdelay3over = frectdelay3over+ochulgo.flist(i).fdelay3over %>
	</td>
	<td>
		<%= round(ochulgo.flist(i).frectdaychulgo,1) %>%<% frectrectdaychulgo = frectrectdaychulgo+ochulgo.flist(i).frectdaychulgo %>
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>
		�Ѱ�
	</td>
	<td><%= CurrFormat(frectbaljutotalno) %></td>		<!--��������ðǼ�-->
	<td><% frectbaesongtotal = (frectcentertotalno / frectbaljutotalno)*100 %> <%= round(frectbaesongtotal,1) %>%	<!--��ü��ۺ���-->
	<td><%= CurrFormat(frectcentertotalno) %></td>		<!--�ѰǼ�-->
	<td><%= frectcancelno %></td>			<!--��ҰǼ�-->
	<td><%= CurrFormat(frecttotalchulgono) %></td>		<!--���Ǽ�-->
	<td><font color="red"><%= CurrFormat(frectdelay0chulgo) %></font></td>		<!--�������-->
	<td><%= frectdelay1chulgo %></td>		<!--1������-->
	<td><%= frectdelay2chulgo %></td>		<!--2������-->
	<td><font color="red"><%= frectdelay3over %></font></td>			<!--3������-->
	<td><% frectrectdaychulgo = (frectdelay0chulgo/frecttotalchulgono)*100 %><%= round(frectrectdaychulgo,1) %>%</td>	<!--���������-->
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=5>���Ǽ� ��� ����</td>
	<td>100%</td>
	<td><% ffrectdelay0chulgo = (frectdelay0chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay0chulgo,1) %>%</td>
	<td><% ffrectdelay1chulgo = (frectdelay1chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay1chulgo,1) %>%</td>
	<td><% ffrectdelay2chulgo = (frectdelay2chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay2chulgo,1) %>%</td>
	<td><% ffrectdelay3chulgo = (frectdelay3over/frecttotalchulgono)*100 %><font color="red"><%= round(ffrectdelay3chulgo,1) %>%</font></td>
	<td></td>
</tr>
</table>
<!-- �Ϻ� �����Ȳ �� -->
<br>
<!-- ���� ��� ���� ����� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td  bgcolor="F4F4F4" width=18%>
	���� ��� ���� �����
	</td>
	<td colspan=8 bgcolor="ffffff">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ݱ�</td>
	<td>1��</td>
	<td>2��</td>
	<td>3��</td>
	<td>4��</td>
	<td>5��</td>
	<td>6��</td>
	<td>�����Ѱ�</td>
	<td>���</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>����ü������Ǽ�</td>
	<td><%= CurrFormat(frectmonthcentertotalno("01")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("02")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("03")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("04")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("05")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("06")) %></td>
	<% dim frectmonthtotalchulgo,frectmonthdangilchulgo,frectdangilper %>
	<td><% frectmonthtotalchulgo = frectmonthcentertotalno("01")+frectmonthcentertotalno("02")+frectmonthcentertotalno("03")+frectmonthcentertotalno("04")+frectmonthcentertotalno("05")+frectmonthcentertotalno("06") %>
	<%= CurrFormat(frectmonthtotalchulgo) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>�������Ǽ�</td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("01")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("02")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("03")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("04")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("05")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("06")) %></td>
	<td>
		<% frectmonthdangilchulgo = frectmonthdelay0chulgo("01")+frectmonthdelay0chulgo("02")+frectmonthdelay0chulgo("03")+frectmonthdelay0chulgo("04")+frectmonthdelay0chulgo("05")+frectmonthdelay0chulgo("06") %>
		<%= CurrFormat(frectmonthdangilchulgo) %>
	</td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>���������</td>
	<td>
		<%
		if frectmonthdelay0chulgo("01") = 0 then
			 frectdangilchulgo1 = 0
		else
			frectdangilchulgo1 = (frectmonthdelay0chulgo("01")/frectmonthcentertotalno("01"))*100
		end if
		response.write  round(frectdangilchulgo1,1) &"%"
		%>
	</td>
	<td>
		<%
		if frectmonthdelay0chulgo("02") = 0 then
			 frectdangilchulgo2 = 0
		else
			frectdangilchulgo2 = (frectmonthdelay0chulgo("02")/frectmonthcentertotalno("02"))*100
		end if
		response.write round(frectdangilchulgo2,1) &"%"
		%>
	</td>
	<td>
		<%
		if frectmonthdelay0chulgo("03") = 0 then
			 frectdangilchulgo3 = 0
		else
			frectdangilchulgo3 = (frectmonthdelay0chulgo("03")/frectmonthcentertotalno("03"))*100
		end if
		response.write round(frectdangilchulgo3,1) &"%"
		%>
	</td>
	<td>
		<%
		if frectmonthdelay0chulgo("04") = 0 then
			 frectdangilchulgo4 = 0
		else
			frectdangilchulgo4 = (frectmonthdelay0chulgo("04")/frectmonthcentertotalno("04"))*100
		end if
		response.write round(frectdangilchulgo4,1) &"%"
		%>
	</td>
	<td>
		<%
		if frectmonthdelay0chulgo("05") = 0 then
			 frectdangilchulgo5 = 0
		else
			frectdangilchulgo5 = (frectmonthdelay0chulgo("05")/frectmonthcentertotalno("05"))*100
		end if
		response.write round(frectdangilchulgo5,1) &"%"
		%>
	</td>
	<td>
		<%
		if frectmonthdelay0chulgo("06") = 0 then
			 frectdangilchulgo6 = 0
		else
			frectdangilchulgo6 = (frectmonthdelay0chulgo("06")/frectmonthcentertotalno("06"))*100
		end if
		response.write round(frectdangilchulgo6,1) &"%"
		%>
	</td>
	<td>
		<%
		if frectmonthdangilchulgo = 0 then
			frectdangilper = 0
		else
			frectdangilper = (frectmonthdangilchulgo/frectmonthtotalchulgo)*100
		end if
		response.write round(frectdangilper,1) &"%"
		%>
	<td></td>
</tr>
<%
frectdangilchulgo1 = 0
frectdangilchulgo2 = 0
frectdangilchulgo3 = 0
frectdangilchulgo4 = 0
frectdangilchulgo5 = 0
frectdangilchulgo6 = 0
frectdangilper= 0
frectmonthdangilchulgo=0
frectmonthtotalchulgo=0
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�Ϲݱ�</td>
	<td>7��</td>
	<td>8��</td>
	<td>9��</td>
	<td>10��</td>
	<td>11��</td>
	<td>12��</td>
	<td>�����Ѱ�</td>
	<td>���</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>����ü������Ǽ�</td>
	<td><%= CurrFormat(frectmonthcentertotalno("07")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("08")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("09")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("10")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("11")) %></td>
	<td><%= CurrFormat(frectmonthcentertotalno("12")) %></td>

	<td>
		<% frectmonthtotalchulgo = frectmonthcentertotalno("07")+frectmonthcentertotalno("08")+frectmonthcentertotalno("09")+frectmonthcentertotalno("10")+frectmonthcentertotalno("11")+frectmonthcentertotalno("12") %>
		<%= CurrFormat(frectmonthtotalchulgo) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>�������Ǽ�</td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("07")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("08")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("09")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("10")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("11")) %></td>
	<td><%= CurrFormat(frectmonthdelay0chulgo("12")) %></td>
	<td>
		<% frectmonthdangilchulgo = frectmonthdelay0chulgo("07")+frectmonthdelay0chulgo("08")+frectmonthdelay0chulgo("09")+frectmonthdelay0chulgo("10")+frectmonthdelay0chulgo("11")+frectmonthdelay0chulgo("12") %>
		<%=CurrFormat( frectmonthdangilchulgo) %></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>���������</td>
	<td>
	<% if frectmonthdelay0chulgo("07") = 0 then
		 frectdangilchulgo1 = 0
		else
		frectdangilchulgo1 = (frectmonthdelay0chulgo("07")/frectmonthcentertotalno("07"))*100
	end if %><%= round(frectdangilchulgo1,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("08") = 0 then
		 frectdangilchulgo2 = 0
		else
		frectdangilchulgo2 = (frectmonthdelay0chulgo("08")/frectmonthcentertotalno("08"))*100
	end if %><%= round(frectdangilchulgo2,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("09") = 0 then
		 frectdangilchulgo3 = 0
		else
		frectdangilchulgo3 = (frectmonthdelay0chulgo("09")/frectmonthcentertotalno("09"))*100
	end if %><%= round(frectdangilchulgo3,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("10") = 0 then
		 frectdangilchulgo4 = 0
		else
		frectdangilchulgo4 = (frectmonthdelay0chulgo("10")/frectmonthcentertotalno("10"))*100
	end if %><%= round(frectdangilchulgo4,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("11") = 0 then
		 frectdangilchulgo5 = 0
		else
		frectdangilchulgo5 = (frectmonthdelay0chulgo("11")/frectmonthcentertotalno("11"))*100
	end if %><%= round(frectdangilchulgo5,1) %>%
	</td>
	<td>
	<% if frectmonthdelay0chulgo("12") = 0 then
		 frectdangilchulgo6 = 0
		else
		frectdangilchulgo3 = (frectmonthdelay0chulgo("12")/frectmonthcentertotalno("12"))*100
	end if %><%= round(frectdangilchulgo6,1) %>%
	</td>
	<td><% if frectmonthdangilchulgo = 0 then
		frectdangilper = 0
		else
		frectdangilper = (frectmonthdangilchulgo/frectmonthtotalchulgo)*100
		end if %>
	<%= round(frectdangilper,1) %>%</td>
	<td></td>
</tr>
</table>
<!-- ���� ��� ���� ����� ��-->

<% else %>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#DDDDFF">
	<td align=center bgcolor="#FFFFFF">�˻� ����� �����ϴ�.</td>
</tr>
</table>

<% end if %>

<%
set ochulgo = nothing
set ochulgomonth = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
