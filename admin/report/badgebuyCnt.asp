<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ ���ŰǼ�
' History : 2015.06.18 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->


<%
	Dim defaultdate1, yyyy1, mm1, dd1, yyyy2, mm2, dd2, MemberShipCardDailylist, i, strTemp, strXML, ChartViDi, strDay, strWeb, strMobile, strApp, strWebLen, strMobileLen, strAppLen, strDate, strDateLen, striOs, striOsLen, strAnd, strAndLen
	Dim vbadgeGubun, sqlstr


	defaultdate1 = dateadd("d",-6,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 10�������� �˻�	
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
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)

	vbadgeGubun = request("badgeGubun")

	If vbadgeGubun="" Then
		vbadgeGubun = "salehunter"
	End If

%>



<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" width="350">
		<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<select name="badgeGubun">
			<option value="salehunter" <% If vbadgeGubun="salehunter" Then %>selected<% End If %>>��������</option>
			<option value="earlybird" <% If vbadgeGubun="earlybird" Then %>selected<% End If %>>�󸮹���</option>
			<option value="brandcool" <% If vbadgeGubun="brandcool" Then %>selected<% End If %>>�귣����</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>

</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">- ������ ���ŰǼ�</font>	
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="50%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>������</td>
	<td>
		<% If Trim(vbadgeGubun) = "salehunter" Or Trim(vbadgeGubun)="earlybird" Then %>
			��ǰ�ڵ�
		<% Else %>
			�귣���
		<% End If %>
	</td>
	<td>���ŰǼ�</td>
</tr>

<% If vbadgeGubun <> "" Then %>
	<%

		Select Case Trim(vbadgeGubun)
			Case "salehunter"
				'// ��������
				sqlstr = " select top 100 "
				sqlstr = sqlstr & "       bi.badgeName, ol.itemid as content, COUNT(*) as cnt "
				sqlstr = sqlstr & " from db_log.dbo.tbl_badge_orderdetail_log as ol "
				sqlstr = sqlstr & " join db_my10x10.dbo.tbl_badge_userObtain as bu on ol.userid=bu.userid "
				sqlstr = sqlstr & " join db_my10x10.dbo.tbl_badge_info as bi on bu.badgeIdx=bi.badgeIdx "
				sqlstr = sqlstr & " join db_item.dbo.tbl_item as i on ol.itemid=i.itemid "
				sqlstr = sqlstr & " 	and isusing='Y' and sellyn='Y' "
				sqlstr = sqlstr & " 	and sailyn='Y' "
				sqlstr = sqlstr & " where bu.obtainDate between '"&yyyy1&"-"&mm1&"-"&dd1&"' and '"&yyyy2&"-"&mm2&"-"&dd2&"' "
				sqlstr = sqlstr & " 	  and ol.isSaleItem='Y' "
				sqlstr = sqlstr & "       and bu.badgeIdx=5 "
				sqlstr = sqlstr & " group by bi.badgeName, ol.itemid "
				sqlstr = sqlstr & " order by cnt desc "
			
			Case "earlybird"
				'// �󸮹���
				sqlstr = " select top 100 "
				sqlstr = sqlstr & "             bi.badgeName, ol.itemid as content, COUNT(*) as cnt "
				sqlstr = sqlstr & "       from db_log.dbo.tbl_badge_orderdetail_log as ol "
				sqlstr = sqlstr & "       join db_my10x10.dbo.tbl_badge_userObtain as bu on ol.userid=bu.userid "
				sqlstr = sqlstr & "       join db_my10x10.dbo.tbl_badge_info as bi on bu.badgeIdx=bi.badgeIdx "
				sqlstr = sqlstr & "       join db_item.dbo.tbl_item as i on ol.itemid=i.itemid "
				sqlstr = sqlstr & "       	and datediff(d,i.regdate,GETDATE())<14 "
				sqlstr = sqlstr & "       	and isusing='Y' and sellyn='Y' "
				sqlstr = sqlstr & "       where bu.obtainDate between '"&yyyy1&"-"&mm1&"-"&dd1&"' and '"&yyyy2&"-"&mm2&"-"&dd2&"' "
				sqlstr = sqlstr & "       	  and ol.isNewItem='Y' "
				sqlstr = sqlstr & "             and bu.badgeIdx=3 "
				sqlstr = sqlstr & "       group by bi.badgeName, ol.itemid "
				sqlstr = sqlstr & "       order by cnt desc "

			Case "brandcool"
				'// �귣����
				sqlstr = " select top 100 "
				sqlstr = sqlstr & "                   bi.badgeName, ol.makerid as content, COUNT(*) as cnt "
				sqlstr = sqlstr & "             from db_log.dbo.tbl_badge_orderdetail_log as ol "
				sqlstr = sqlstr & "             join db_my10x10.dbo.tbl_badge_userObtain as bu on ol.userid=bu.userid "
				sqlstr = sqlstr & "             join db_my10x10.dbo.tbl_badge_info as bi on bu.badgeIdx=bi.badgeIdx "
				sqlstr = sqlstr & "             join db_item.dbo.tbl_item as i on ol.itemid=i.itemid "
				sqlstr = sqlstr & "             	and isusing='Y' and sellyn='Y' "
				sqlstr = sqlstr & "             where bu.obtainDate between '"&yyyy1&"-"&mm1&"-"&dd1&"' and '"&yyyy2&"-"&mm2&"-"&dd2&"' "
				sqlstr = sqlstr & "                   and bu.badgeIdx=6 "
				sqlstr = sqlstr & "             group by bi.badgeName, ol.makerid "
				sqlstr = sqlstr & "             order by cnt desc "

			Case Else
				sqlstr = " select top 100 "
				sqlstr = sqlstr & "       bi.badgeName, ol.itemid as content, COUNT(*) as cnt "
				sqlstr = sqlstr & " from db_log.dbo.tbl_badge_orderdetail_log as ol "
				sqlstr = sqlstr & " join db_my10x10.dbo.tbl_badge_userObtain as bu on ol.userid=bu.userid "
				sqlstr = sqlstr & " join db_my10x10.dbo.tbl_badge_info as bi on bu.badgeIdx=bi.badgeIdx "
				sqlstr = sqlstr & " join db_item.dbo.tbl_item as i on ol.itemid=i.itemid "
				sqlstr = sqlstr & " 	and isusing='Y' and sellyn='Y' "
				sqlstr = sqlstr & " 	and sailyn='Y' "
				sqlstr = sqlstr & " where bu.obtainDate between '"&yyyy1&"-"&mm1&"-"&dd1&"' and '"&yyyy2&"-"&mm2&"-"&dd2&"' "
				sqlstr = sqlstr & " 	  and ol.isSaleItem='Y' "
				sqlstr = sqlstr & "       and bu.badgeIdx=5 "
				sqlstr = sqlstr & " group by bi.badgeName, ol.itemid "
				sqlstr = sqlstr & " order by cnt desc "
		End Select
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
	%>
	<% If Not(rsget.bof Or rsget.eof) Then %>
		<% 
			Do Until rsget.eof
		%>
			<tr bgcolor="#FFFFFF" align="center">
				<td><%=rsget("badgename")%></td>
				<td><%=rsget("content")%></td>
				<td><%=rsget("cnt")%></td>
			</tr>
		<%
			rsget.movenext
			Loop
		%>
		</table>
	<% Else %>
		<table width="50%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		<tr align="center" bgcolor="#FFFFFF">
			<td colspan="4">�˻� ����� �����ϴ�.</td>
		</tr>
		</table>
	<%
		End If
		rsget.close
	%>
<% End If %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->