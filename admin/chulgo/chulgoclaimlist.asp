<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  Ŭ���� ����
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
dim yyyy , mm ,frectmonthtotalchulgo,frectmonthdangilchulgo,frectdangilper
dim frectbaljutotalno, frectrectbaesong,frectcentertotalno,frectcancelno,frecttotalchulgono, frectclaimA000chulgo
dim frectclaimA001chulgo,frectclaimA002chulgo,frectclaimsum,frectclaimchulgo,frectbaesongtotal ,frectrectdaychulgo
dim ffrectdelay0chulgo,ffrectdelay1chulgo,ffrectdelay2chulgo,ffrectdelay3chulgo , frectclaimsumtotal
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
	var excel = window.open('/admin/chulgo/chulgoclaimlist_excel.asp?yyyy1=' + yyyy + ' &mm1=' +mm ,'excelsheet','width=1024,height=768,scrollbars=yes,resizable=yes');
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
		��: <% DrawYMBox yyyy,mm %>	
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
<br>		
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">			
	</td>
	<td align="right">				
	</td>
</tr>
</table>
<!-- �׼� �� -->

<% if ochulgo.FTotalCount > 0 then %>
	<!-- �Ϻ� �����Ȳ ���� -->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=2>
			�Ϻ� Ŭ���� ��� ��Ȳ
		</td>
		<td colspan=9>
			��ǥ : ��ü������Ǽ� ��� 1% �̳� <input type="button" onclick="ExcelSheet('<%= yyyy %>','<%= mm %>')" value="������ ���" class="button">
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td rowspan="2">��¥</td>
		<td rowspan="2">�ѹ��ְǼ�</td>
		<td rowspan="2">��ü��ۺ���</td>
		<td colspan=3>��ü��۰Ǽ�</td>
		<td colspan=4>Ŭ���� �����</td>
		<td rowspan="2">Ŭ����������</td>	
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td>�ѰǼ�</td>
		<td>��ҰǼ�</td>
		<td>���Ǽ�</td>
		<td>�±�ȯ���</td>
		<td>������߼�</td>
		<td>���񽺹߼�</td>
		<td>�Ұ�</td>
	</tr>
	<% for i=0 to ochulgo.FTotalCount - 1 %>
	<tr bgcolor="ffffff" align="center">
		<td>		
			<%= ochulgo.flist(i).fmm %>�� <%= ochulgo.flist(i).fdd %>��
		</td>
		<td>		
			<%= CurrFormat(ochulgo.flist(i).fbaljutotalno) %><% frectbaljutotalno = frectbaljutotalno+ochulgo.flist(i).fbaljutotalno %>
		</td>
		<td>		
			<%= round(ochulgo.flist(i).frectbaesong,1) %>%<% frectrectbaesong = frectrectbaesong+ochulgo.flist(i).frectbaesong %>
		</td>
		<td>		
			<%= CurrFormat(ochulgo.flist(i).fcentertotalno)	%><% frectcentertotalno = frectcentertotalno+ochulgo.flist(i).fcentertotalno %>
		</td>
		<td>		
			<%= ochulgo.flist(i).fcancelno %><% frectcancelno = frectcancelno+ochulgo.flist(i).fcancelno %>
		</td>
		<td>		
			<%= CurrFormat(ochulgo.flist(i).ftotalchulgono) %><% frecttotalchulgono = frecttotalchulgono+ochulgo.flist(i).ftotalchulgono %>
		</td>
		<td>		
			<%= ochulgo.flist(i).fclaimA000 %><% frectclaimA000chulgo = frectclaimA000chulgo+ochulgo.flist(i).fclaimA000 %>
		</td>
		<td>		
			<%= ochulgo.flist(i).fclaimA001 %><% frectclaimA001chulgo = frectclaimA001chulgo+ochulgo.flist(i).fclaimA001 %>
		</td>
		<td>		
			<%= ochulgo.flist(i).fclaimA002 %><% frectclaimA002chulgo = frectclaimA002chulgo+ochulgo.flist(i).fclaimA002 %>
		</td>
		<td>		
			<% frectclaimsum = ochulgo.flist(i).fclaimA000+ochulgo.flist(i).fclaimA001+ochulgo.flist(i).fclaimA002 %><%= frectclaimsum %>
		</td>
		<td>			
			<% if ochulgo.flist(i).ftotalchulgono <> 0 then %>				
				<% frectclaimchulgo = (ochulgo.flist(i).fclaimA000/ochulgo.flist(i).ftotalchulgono)*100 %>
			<% else %>
				<% frectclaimchulgo = 0 %>
			<% end if %>
			<%= round(frectclaimchulgo,1) %>%
		</td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�Ѱ�</td>
		<td><%= CurrFormat(frectbaljutotalno) %></td>		<!--�ѹ��ְǼ�-->
		<td><% frectbaesongtotal = (frectcentertotalno / frectbaljutotalno)*100 %> <%= round(frectbaesongtotal,1) %>%	<!--��ü��ۺ���-->
		<td><%= CurrFormat(frectcentertotalno) %></td>		<!--�ѰǼ�-->
		<td><%= frectcancelno %></td>			<!--��ҰǼ�-->
		<td><%= CurrFormat(frecttotalchulgono) %></td>		<!--���Ǽ�-->
		<td><%= CurrFormat(frectclaimA000chulgo) %></td>	<!--�±�ȯ���-->
		<td><%= CurrFormat(frectclaimA001chulgo) %></td>		<!--������߼�-->
		<td><%= CurrFormat(frectclaimA002chulgo) %></td>		<!--���񽺹߼�-->
		<td><% frectclaimsumtotal = frectclaimA000chulgo+frectclaimA001chulgo+frectclaimA002chulgo %><%= CurrFormat(frectclaimsumtotal) %></td>			<!--�Ұ�-->
		<td><% frectrectdaychulgo = (frectclaimsumtotal/frecttotalchulgono)*100 %><%= round(frectrectdaychulgo,1) %>%</td>	<!--Ŭ����������-->
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=5>���Ǽ� ��� ����</td>
		<td>100%</td>
		<td><% ffrectdelay0chulgo = (frectclaimA000chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay0chulgo,1) %>%</td>
		<td><% ffrectdelay1chulgo = (frectclaimA001chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay1chulgo,1) %>%</td>
		<td><% ffrectdelay2chulgo = (frectclaimA002chulgo/frecttotalchulgono)*100 %><%= round(ffrectdelay2chulgo,1) %>%</td>
		<td><% ffrectdelay3chulgo = (frectclaimsumtotal/frecttotalchulgono)*100 %><%= round(ffrectdelay3chulgo,1) %>%</td>
		<td></td>
	</tr>
	</table>
	<!-- �Ϻ� �����Ȳ �� -->
	<br>
	<!-- ���� ��� ���� ����� ����-->	
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr>
		<td  bgcolor="F4F4F4" width=18%>
		���� Ŭ���� ������
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
		<td>
			<% frectmonthtotalchulgo = frectmonthcentertotalno("01")+frectmonthcentertotalno("02")+frectmonthcentertotalno("03")+frectmonthcentertotalno("04")+frectmonthcentertotalno("05")+frectmonthcentertotalno("06") %>
			<%= CurrFormat(frectmonthtotalchulgo) %>
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>Ŭ�������Ǽ�</td>
		<td><%= CurrFormat(frectmonthclaimchulgo("01")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("02")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("03")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("04")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("05")) %></td>
		<td><%= CurrFormat( frectmonthclaimchulgo("06")) %></td>
		<td>
			<% frectmonthdangilchulgo = frectmonthclaimchulgo("01")+frectmonthclaimchulgo("02")+frectmonthclaimchulgo("03")+frectmonthclaimchulgo("04")+frectmonthclaimchulgo("05")+frectmonthclaimchulgo("06") %>
			<%= CurrFormat(frectmonthdangilchulgo) %>
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>Ŭ����������</td>	
		<td>
			<% if frectmonthclaimchulgo("01") <> 0 then %>
				<% frectdangilchulgo1 = (frectmonthclaimchulgo("01")/frectmonthcentertotalno("01"))*100 %>
			<% else %>
				<% frectdangilchulgo1 = 0 %>
			<% end if %>
			<%= round(frectdangilchulgo1,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("02") <> 0 then %>
			<% frectdangilchulgo2 = (frectmonthclaimchulgo("02")/frectmonthcentertotalno("02"))*100 %>
			<% else %>
				<% frectdangilchulgo2 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo2,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("03") <> 0 then %>
				<% frectdangilchulgo3 = (frectmonthclaimchulgo("03")/frectmonthcentertotalno("03"))*100 %>
			<% else %>
				<% frectdangilchulgo3 = 0 %>
			<% end if%>
			<%= round(frectdangilchulgo3,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("04") <>0 then %>
				<% frectdangilchulgo4 = (frectmonthclaimchulgo("04")/frectmonthcentertotalno("04"))*100 %>
			<% else %>
				<% frectdangilchulgo4 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo4,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("05") <> 0 then %>
				<% frectdangilchulgo5 = (frectmonthclaimchulgo("05")/frectmonthcentertotalno("05"))*100 %>
			<% else %>
				<%frectdangilchulgo5 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo5,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("06") <> 0 then %>
				<% frectdangilchulgo6 = (frectmonthclaimchulgo("06")/frectmonthcentertotalno("06"))*100 %>
			<% else %>
				<% frectdangilchulgo6 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo6,1) %>%
		</td>
		<td>
			<% frectdangilper = (frectdangilchulgo1+frectdangilchulgo2+frectdangilchulgo3+frectdangilchulgo4+frectdangilchulgo5+frectdangilchulgo6)/6 %>
			<%= round(frectdangilper,1) %>%
		</td>
		<td>��ǥ : 1% �̳�</td>
	</tr>
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
			<%= CurrFormat(frectmonthtotalchulgo) %>
		</td> 
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>Ŭ�������Ǽ�</td>
		<td><%= CurrFormat(frectmonthclaimchulgo("07")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("08")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("09")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("10")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("11")) %></td>
		<td><%= CurrFormat(frectmonthclaimchulgo("12")) %></td>
		<td>
			<% frectmonthdangilchulgo = frectmonthclaimchulgo("07")+frectmonthclaimchulgo("08")+frectmonthclaimchulgo("09")+frectmonthclaimchulgo("10")+frectmonthclaimchulgo("11")+frectmonthclaimchulgo("12") %>
			<%= CurrFormat(frectmonthdangilchulgo) %>
		</td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>Ŭ����������</td>
		<td>
			<% if frectmonthclaimchulgo("07") <> 0 then %>
				<% frectdangilchulgo1 = (frectmonthclaimchulgo("07")/frectmonthcentertotalno("07"))*100 %>
			<% else %>
				<% frectdangilchulgo1 = 0 %>
			<% end if %>
			<%= round(frectdangilchulgo1,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("08") <> 0 then %>
			<% frectdangilchulgo2 = (frectmonthclaimchulgo("08")/frectmonthcentertotalno("08"))*100 %>
			<% else %>
				<% frectdangilchulgo2 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo2,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("09") <> 0 then %>
				<% frectdangilchulgo3 = (frectmonthclaimchulgo("09")/frectmonthcentertotalno("09"))*100 %>
			<% else %>
				<% frectdangilchulgo3 = 0 %>
			<% end if%>
			<%= round(frectdangilchulgo3,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("10") <>0 then %>
				<% frectdangilchulgo4 = (frectmonthclaimchulgo("10")/frectmonthcentertotalno("10"))*100 %>
			<% else %>
				<% frectdangilchulgo4 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo4,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("11") <> 0 then %>
				<% frectdangilchulgo5 = (frectmonthclaimchulgo("11")/frectmonthcentertotalno("11"))*100 %>
			<% else %>
				<%frectdangilchulgo5 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo5,1) %>%
		</td>
		<td>
			<% if frectmonthclaimchulgo("12") <> 0 then %>
				<% frectdangilchulgo6 = (frectmonthclaimchulgo("12")/frectmonthcentertotalno("12"))*100 %>
			<% else %>
				<% frectdangilchulgo6 = 0 %>
			<% end if %>	
			<%= round(frectdangilchulgo6,1) %>%
		</td>
		<td>
			<% frectdangilper = (frectdangilchulgo1+frectdangilchulgo2+frectdangilchulgo3+frectdangilchulgo4+frectdangilchulgo5+frectdangilchulgo6)/6 %>
			<%= round(frectdangilper,1) %>%
		</td>
		<td></td>
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
set ochulgo = nothing
set ochulgomonth = nothing
%>	

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->