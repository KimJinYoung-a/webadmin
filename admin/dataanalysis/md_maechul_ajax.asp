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
dim startdate, enddate, i
	startdate = Request("startdate")
	enddate = Request("enddate")

if startdate="" then startdate=date
if enddate="" then enddate=date

dim omaechul
set omaechul = new cdataanalysis
	omaechul.FPageSize = 10
	omaechul.FCurrPage = 1
	omaechul.frectstartdate = startdate
	omaechul.frectenddate = enddate
	omaechul.Getmaechullist
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="6">
		<b>��ǥ����</b>
		&nbsp;&nbsp;�� ��ǥ(30�� ���� ������), ����(����ϱ���, ��ȯ&��ǰ ����, 30�� ���� ������)
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="6">
	    ��ȸ�Ⱓ : <%= startdate %>~<%= enddate %>
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

<% if omaechul.FResultCount > 0 then %>
	<% for i = 0 to omaechul.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			<%= getgubunname(omaechul.FItemList(i).fgubun) %>
		</td>
		<td align="right">
			<%= CurrFormat(omaechul.FItemList(i).fmaechul) %>
		</td>
		<td align="right">
			<%= CurrFormat(omaechul.FItemList(i).fprofit) %>
		</td>
		<td align="right">
			<%
			'/��ǥ
			if omaechul.FItemList(i).fgubun="purpose" then 
			%>
				<% if omaechul.fcurrentmaechul<>0 and omaechul.fpurposemaechul<>0 then %>
					<%= getgrade(round((omaechul.fcurrentmaechul/omaechul.fpurposemaechul)*100,2)) %>
				<% else %>
					<img src='/images/grade/grade_90DOWN.png'>
				<% end if %>
			<%
			'/����
			elseif omaechul.FItemList(i).fgubun="currentmaechul" then 
			%>
				<% if omaechul.fcurrentmaechul<>0 and omaechul.fpurposemaechul<>0 then %>
					<%= round((omaechul.fcurrentmaechul/omaechul.fpurposemaechul)*100,2) %>%
				<% else %>
					0%
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<%
			'/��ǥ
			if omaechul.FItemList(i).fgubun="purpose" then 
			%>
				<% if omaechul.fcurrentprofit<>0 and omaechul.fpurposeprofit<>0 then %>
					<%= getgrade(round((omaechul.fcurrentprofit/omaechul.fpurposeprofit)*100,2)) %>
				<% else %>
					<img src='/images/grade/grade_90DOWN.png'>
				<% end if %>
			<%
			'/����
			elseif omaechul.FItemList(i).fgubun="currentmaechul" then 
			%>
				<% if omaechul.fcurrentprofit<>0 and omaechul.fpurposeprofit<>0 then %>
					<%= round((omaechul.fcurrentprofit/omaechul.fpurposeprofit)*100,2) %>%
				<% else %>
					0%
				<% end if %>
			<% end if %>
		</td>
		<td align="right">
			<%
			'/������ ���� ������
			if omaechul.FItemList(i).fgubun="currentmaechul" then
				'if dateserial(calyyyy, Format00(2,calmm), "01") < dateserial(Year(date), Format00(2,Month(date)), "01") then
			%>
					<% if omaechul.fcurrentmaechul<>0 and omaechul.fbeforemaechul<>0 then %>
						<%= round((((omaechul.fcurrentmaechul/omaechul.fbeforemaechul)*100) -100),2) %>%
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

<% set omaechul = nothing %>
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->