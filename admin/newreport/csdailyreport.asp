<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/csdailyreportcls.asp" -->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2,i

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

dim temp

if (yyyy1="") then
	
	temp=dateadd("y",now(),-7)
	temp=split(temp,"-")
	
	yyyy1=temp(0)
	mm1=temp(1)
	dd1=temp(2)

end if

if (yyyy2="") then
	
	temp=dateadd("y",now(),0)
	temp=split(temp,"-")
	
	yyyy2=temp(0)
	mm2=temp(1)
	dd2=temp(2)
	
end if





dim qna
set qna = new CsTotal
qna.yyyy1=yyyy1
qna.mm1=mm1
qna.dd1=dd1
qna.yyyy2=yyyy2
qna.mm2=mm2
qna.dd2=dd2
qna.GetCsTotal

%>
<table width="100%" border="0" cellpadding="0" cellspacing="0"class="a">
	<tr>
		<td align=left>�����Ǽ� :	<img src="/images/dot1.gif" height="4" width="20"> �����Ǽ� : <img src="/images/dot2.gif" height="4" width="20">  ��� �亯�ð� : <img src="/images/dot4.gif" height="4" width="20"></td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
		<input type="hidden" name="showtype" value="showtype">
		<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		�˻��Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>

		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>
<table width="100%" height=50 border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td width="10%" align=center>��¥</td>
	<!--<td width="20%" align=center>�������Ǽ�</td>-->
	<!--<td width="20%" align=center>��ó���Ǽ�</td>-->
	<td width="90%" align=center>����</td>
	<!--<td width="20%" align=center>��մ亯�ð�</td>-->
</tr>
<% if qna.FTotalCount < 1  then %>
<% else %>
<% For i=0 to qna.FTotalCount-1 %>

<tr bgcolor="#FFFFFF">
	<td width="7%" align=center><%= qna.Items(i).Fday %></td>
	<td>
		<table width="90%" height=50 border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td align=left>
					<img src="/images/dot1.gif" height="4" width="<%= qna.Items(i).FRegcnt/qna.maxregcnt*90 %>%"><%= qna.Items(i).FRegcnt %><br>
					<img src="/images/dot2.gif" height="4" width="<%= qna.Items(i).FDelaycnt/qna.maxregcnt*90 %>%"><%= qna.Items(i).FDelaycnt%><br>
				<div align="left">
					<img src="/images/dot4.gif" height="4" width="<%= left(qna.Items(i).FAvgtime,5)/40*90 %>%"><%= left(qna.Items(i).FAvgtime,5)%></div>
				</td>
			</tr>
		</table>
	</td>
	
<% next %>
<% end if %>
</tr>
<tr bgcolor=#FFFFFF>
	<td width="7%" align=center>��&nbsp;&nbsp;��</td>
	<td align=center width="100%">
	<table  width="100%" border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
			<td width="25%" align=center>�� ���� : <%= qna.regtotal %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="25%" align=center>ó�� �Ǽ� �հ� : <%= qna.fintotal %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="25%" align=center>���� �Ǽ� �հ� : <%= qna.delaytotal %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="25%" align=center>��� �亯 �ð�: <%= left(qna.avgtotal,5) %></td>
		</tr>
	</table>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->