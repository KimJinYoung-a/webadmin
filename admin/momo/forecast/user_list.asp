<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.11.19 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oforecast,i,page , cardidx , isusing , yyyy , mm , owinner , winner0 , winner1 , winner2 , contents0, contents1, contents2
	menupos = request("menupos")
	yyyy = request("yyyy1")
	mm = request("mm1")
	
	if yyyy = "" then yyyy = year(date())
	if mm = "" then mm = month(date())	

'// ����Ʈ
set oforecast = new cforecast_list
	oforecast.frectyyyymm = yyyy & "-" & mm
	oforecast.fuser_list()

'// ����Ʈ
set owinner = new cforecast_list
	owinner.frectyyyymm = yyyy & "-" & mm
	owinner.frectgubun = "0"
	owinner.fuser_winner()

	if owinner.ftotalcount > 0 then	
		for i = 0 to owinner.ftotalcount - 1
			if owinner.FItemList(i).forderno = "0" then
				winner0 = owinner.FItemList(i).fuserid
				contents0 = owinner.FItemList(i).fcontents
			elseif owinner.FItemList(i).forderno = "1" then
				winner1 = owinner.FItemList(i).fuserid
				contents1 = owinner.FItemList(i).fcontents
			elseif owinner.FItemList(i).forderno = "2" then
				winner2 = owinner.FItemList(i).fuserid
				contents2 = owinner.FItemList(i).fcontents												
			end if
		next
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
		if (winnerfrm.contents0.value==''){
			alert('1�� �������� �Է����ּ���');
			winnerfrm.contents0.focus();
			return;
		}
		if (winnerfrm.winner1.value==''){
			alert('2�� ���̵� �Է����ּ���');
			winnerfrm.winner2.focus();
			return;
		}
		if (winnerfrm.contents1.value==''){
			alert('2�� �������� �Է����ּ���');
			winnerfrm.contents2.focus();
			return;
		}
		if (winnerfrm.winner2.value==''){
			alert('3�� ���̵� �Է����ּ���');
			winnerfrm.winner3.focus();
			return;
		}
		if (winnerfrm.contents2.value==''){
			alert('3�� �������� �Է����ּ���');
			winnerfrm.contents3.focus();
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
<form name="winnerfrm" action="user_process.asp" method="post">
<input type="hidden" name="mode" value="winneredit">
<input type="hidden" name="yyyy" value="<%=yyyy%>">
<input type="hidden" name="mm" value="<%=mm%>">
<tr>
	<td align="left">				
		��<%=yyyy%>�� <%=MM%>�� ��÷��<Br>
		1�� : ���̵�<input type="text" name="winner0" value="<%=winner0%>"> &nbsp;&nbsp;&nbsp;������<input type="text" name="contents0" value="<%=contents0%>"><br>
		2�� : ���̵�<input type="text" name="winner1" value="<%=winner1%>"> &nbsp;&nbsp;&nbsp;������<input type="text" name="contents1" value="<%=contents1%>"><br>
		3�� : ���̵�<input type="text" name="winner2" value="<%=winner2%>"> &nbsp;&nbsp;&nbsp;������<input type="text" name="contents2" value="<%=contents2%>">
		<input type="button" onclick="winnerreg(<%=yyyy%>-<%=mm%>);" class="button" value="����"><br>		
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<br>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oforecast.FTotalCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oforecast.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>	
	<td align="center">��</td>
	<td align="center">�����ϼ�</td>	
	<td align="center">���</td>
</tr>
<% for i=0 to oforecast.ftotalcount -1 %>			

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<td align="center">
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fuserid %>
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fusercount %> [<%=fix((oforecast.FItemList(i).fusercount / datediff("d",DateSerial(yyyy, mm,1),DateSerial(yyyy, mm+1,1))) * 100)%> %]		
	</td>
	<td align="center">
	</td>			
</tr>   

<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oforecast.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oforecast.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oforecast.StartScrollPage to oforecast.StartScrollPage + oforecast.FScrollCount - 1 %>
			<% if (i > oforecast.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oforecast.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oforecast.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oforecast = nothing
	set owinner = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->