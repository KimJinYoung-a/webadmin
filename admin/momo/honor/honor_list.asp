<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ��������
' Hieditor : 2010.11.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim i, yyyy, mm,owinner , winner0 , winner1 , winner2 , winner3, winner4, winner5
	menupos = request("menupos")
	yyyy = request("yyyy1")
	mm = request("mm1")
	
	if yyyy = "" then yyyy = year(date())
	if mm = "" then mm = month(date())	

'//��÷�� ����
set owinner = new chonor_list
	owinner.frectyyyymm = yyyy & "-" & mm
	owinner.frectgubun = "1"
	owinner.fhonor_winner()

	if owinner.ftotalcount = 6 then
		winner0 = owinner.FItemList(0).fuserid
		winner1 = owinner.FItemList(1).fuserid
		winner2 = owinner.FItemList(2).fuserid
		winner3 = owinner.FItemList(3).fuserid
		winner4 = owinner.FItemList(4).fuserid
		winner5 = owinner.FItemList(5).fuserid
	end if
%>

<script language="javascript">

	//��÷�� ���
	function winnerreg(){
		if (winnerfrm.winner0.value==''){
			alert('1�� ���̵� �Է����ּ���');
			winnerfrm.winner0.focus();
			return;
		}
		if (winnerfrm.winner1.value==''){
			alert('2�� ���̵� �Է����ּ���');
			winnerfrm.winner1.focus();
			return;
		}
		if (winnerfrm.winner2.value==''){
			alert('2�� ���̵� �Է����ּ���');
			winnerfrm.winner2.focus();
			return;
		}
		if (winnerfrm.winner3.value==''){
			alert('3�� ���̵� �Է����ּ���');
			winnerfrm.winner3.focus();
			return;
		}
		if (winnerfrm.winner4.value==''){
			alert('3�� ���̵� �Է����ּ���');
			winnerfrm.winner4.focus();
			return;
		}
		if (winnerfrm.winner5.value==''){
			alert('3�� ���̵� �Է����ּ���');
			winnerfrm.winner5.focus();
			return;
		}														
		winnerfrm.submit()
	}
	
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="cardidx">	
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>		
	<td align="left">
		��¥ : <% DrawYMBox yyyy , mm %> 	
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
<form name="winnerfrm" action="honor_process.asp" method="post">
<input type="hidden" name="mode" value="winneredit">
<input type="hidden" name="yyyy" value="<%=yyyy%>">
<input type="hidden" name="mm" value="<%=TwoNumber(mm)%>">
<tr>
	<td align="left">				
		��<%=yyyy%>�� <%=MM%>�� ��÷��<Br>
		1�� : ���̵�<input type="text" name="winner0" value="<%=winner0%>"><br>
		2�� : ���̵�<input type="text" name="winner1" value="<%=winner1%>"><input type="text" name="winner2" value="<%=winner2%>"><br>
		3�� : ���̵�<input type="text" name="winner3" value="<%=winner3%>"><input type="text" name="winner4" value="<%=winner4%>"><input type="text" name="winner5" value="<%=winner5%>">
		<input type="button" onclick="winnerreg(<%=yyyy%>-<%=mm%>);" class="button" value="����"><br>		
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<%
	set owinner = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->