<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<%
dim oVacation
dim empno, arrList, intLoop
dim sumMon, sumYear
empno =  requestCheckVar(Request("empno"),32)
sumMon = 0
sumYear = 0
if empno = "" then
	Alert_return("����� �������� �ʽ��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���")
	response.end
end if

Set oVacation = new CTenByTenVacation
	oVacation.Fempno = empno
	arrList = oVacation.fnGetPartList
Set oVacation = nothing
%>
</head>
<body leftmargin="10" topmargin="10">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>�������� �ް��ð� ��������</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td> <div style="padding:10px;">���: <%=empno%>&nbsp;&nbsp;<%IF isArray(arrList) THEN%>�̸�: <%=arrList(4,0)%>&nbsp;&nbsp;�Ի���: <%=arrList(5,0)%><%end if%></div>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">��¥</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�� �����ϼ�</td>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�� �����ϼ�</td>
		</tr>
		<%IF isArray(arrList) THEN
			For intLoop = 0 to UBound(arrList,2)

			%>
		<tr>
			<td  bgcolor="#FFFFFF"  align="center"><%=arrList(1,intloop)%></td>
			<td  bgcolor="#FFFFFF" align="center"><%if arrList(6,intLoop) = 13 then%><%=Int((arrList(2,intloop)/60)+0.5)%><%else%><%=arrList(2,intloop)%><%end if%></td>
			<td  bgcolor="#FFFFFF" align="center"><%if arrList(6,intLoop) = 13 then%><%=CeilValue((arrList(3,intloop)/60))%><%else%><%=arrList(3,intLoop)%><%end if%></td>
		</tr>
		<%
				sumMon = sumMon + arrList(2,intloop)
				sumYear = sumYear + arrList(3,intloop)
			Next
	%>
		<tr>
			<td bgcolor="#e3f1fb" align="center">���� (�Ҽ��� �հ� �ø�ó��)</td>
			<td bgcolor="#e3f1fb" align="center"><%if arrList(6,0) = 13 then%><%=CeilValue(sumMon/60)%><%else%><%=sumMon%><%end if%></td>
			<td bgcolor="#e3f1fb" align="center"><%if arrList(6,0) = 13 then%><%=CeilValue(sumYear/60)%><%else%><%=sumYear%><%end if%></td>
		</tr>
			<%END IF%>
	</td>
</tr>
</table>
</body>
</html>