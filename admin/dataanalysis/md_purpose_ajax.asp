<%@ language=vbscript %>
<% option explicit %>
<% Response.CharSet = "EUC-KR" %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �����ͺм�
' History : 2016.01.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_salesissue_cls.asp"-->
<%
dim calyyyy, calmm
	calyyyy = Request("calyyyy")
	calmm = Request("calmm")

if calyyyy="" then calyyyy=Year(now)
if calmm="" then calmm=Month(now)

dim opurpose, i
set opurpose = new cdataanalysis
	opurpose.FPageSize = 10
	opurpose.FCurrPage = 1
	opurpose.frectyyyy = calyyyy
	opurpose.frectmm = Format00(2,calmm)
	opurpose.Getpurposelist
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="6">
		<b>��ǥ����</b>
		&nbsp;&nbsp;�� ��ǥ(30�� ���� ������), ����(����ϱ���, ��ȯ&��ǰ ����, 6�ð� ���� ������)
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="6">
	    <input type="button" value="��" onclick="gopurpose('<%= calyyyy-1 %>','<%= calmm %>');" class="calBtn">
	    <b><%=calyyyy%></b>
	    <input type="button" value="��" onclick="gopurpose('<%= calyyyy+1 %>','<%= calmm %>');" class="calBtn">
	    &nbsp;/&nbsp;
	    <input type="button" value="��" onclick="gopurpose('<%= chkIIF(calmm-1<1,calyyyy-1,calyyyy) %>','<%= chkIIF(calmm-1<1,"12",calmm-1) %>');" class="calBtn">
	    <b><%=calmm%></b>
	    <input type="button" value="��" onclick="gopurpose('<%= chkIIF(calmm+1>12,calyyyy+1,calyyyy) %>','<%= chkIIF(calmm+1>12,"1",calmm+1) %>');" class="calBtn">
	    
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td></td>
	<td>����</td>
	<td>����</td>
	<td>���� �޼���</td>
	<td>���� �޼���</td>
	<td>������ ���� ������</td>
</tr>

<% if opurpose.FResultCount > 0 then %>
	<% for i = 0 to opurpose.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			<%= getgubunname(opurpose.FItemList(i).fgubun) %>
		</td>
		<td align="right">
			<%= CurrFormat(opurpose.FItemList(i).fmaechul) %>
		</td>
		<td align="right">
			<%= CurrFormat(opurpose.FItemList(i).fprofit) %>
		</td>
		<td align="right">
			<%
			'/��ǥ
			if opurpose.FItemList(i).fgubun="purpose" then 
			%>
				<% if opurpose.fcurrentmaechul<>0 and opurpose.fpurposemaechul<>0 then %>
					<%= getgrade(round((opurpose.fcurrentmaechul/opurpose.fpurposemaechul)*100,2)) %>
				<% else %>
					<img src='/images/grade/grade_90DOWN.png'>
				<% end if %>
			<%
			'/����
			elseif opurpose.FItemList(i).fgubun="currentmaechul" then 
			%>
				<% if opurpose.fcurrentmaechul<>0 and opurpose.fpurposemaechul<>0 then %>
					<%= round((opurpose.fcurrentmaechul/opurpose.fpurposemaechul)*100,2) %>%
				<% else %>
					0%
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<%
			'/��ǥ
			if opurpose.FItemList(i).fgubun="purpose" then 
			%>
				<% if opurpose.fcurrentprofit<>0 and opurpose.fpurposeprofit<>0 then %>
					<%= getgrade(round((opurpose.fcurrentprofit/opurpose.fpurposeprofit)*100,2)) %>
				<% else %>
					<img src='/images/grade/grade_90DOWN.png'>
				<% end if %>
			<%
			'/����
			elseif opurpose.FItemList(i).fgubun="currentmaechul" then 
			%>
				<% if opurpose.fcurrentprofit<>0 and opurpose.fpurposeprofit<>0 then %>
					<%= round((opurpose.fcurrentprofit/opurpose.fpurposeprofit)*100,2) %>%
				<% else %>
					0%
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<%
			'/������ ���� ������
			if opurpose.FItemList(i).fgubun="currentmaechul" then
				'if dateserial(calyyyy, Format00(2,calmm), "01") < dateserial(Year(date), Format00(2,Month(date)), "01") then
			%>
					<% if opurpose.fcurrentmaechul<>0 and opurpose.fbeforemaechul<>0 then %>
						<%= round((((opurpose.fcurrentmaechul/opurpose.fbeforemaechul)*100) -100),2) %>%
					<% else %>
						0%
					<% end if %>
				<% 'else %>
					<!--������-->
				<% 'end if %>
			<% end if %>
		</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center" class="page_link">��ϵ� ��ǥ�� �����ϴ�.</td>
	</tr>
<% end if %>
</table>

<% set opurpose = nothing %>
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->