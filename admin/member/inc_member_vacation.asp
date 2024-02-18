<%
	'// 직원 휴가신청 목록
	Dim vacationSql, addSql

	'팀장이상이거나, 시스템팀이면 전팀 표시 / 파트장은 자기팀만
	if Not((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or session("ssAdminPsn")=7) then
		addSql = " and p.part_sn='" & session("ssAdminPsn") & "'"
	end if

	vacationSql = "select " &_
				" 	p.part_sn  " &_
				" 	, dv.departmentNameFull as part_name	 " &_
				" 	, sum(vm.requestedday) as totrequestedday " &_
				" 	, count(vm.requestedday) as cntrequestedday " &_
				" from " &_
				" 	[db_partner].[dbo].tbl_vacation_master vm " &_
				" 	join [db_partner].[dbo].tbl_user_tenbyten p " &_
				" 	on " &_
				" 		vm.empno = p.empno " &_
				" 	join [db_partner].dbo.tbl_partInfo pa " &_
				" 	on " &_
				" 		p.part_sn = pa.part_sn " &_
				" 	left join db_partner.dbo.vw_user_department dv on p.department_id = dv.cid " &_
				" where 1 = 1 " &_
				" 	and p.isusing = 1" &_
				" and (p.statediv ='Y' or (p.statediv ='N' and datediff(dd,p.retireday,getdate())<=0))" &_
				" 	and vm.requestedday > 0 " &_
				" 	and vm.endday >= getdate() " &_
				" 	and vm.deleteyn <> 'Y' " & addSql &_
				" group by " &_
				" 	p.part_sn, pa.part_name, dv.departmentNameFull " &_
				" order by " &_
				" 	dv.departmentNameFull "

	'response.write vacationSql & "<br>"
	rsget.Open vacationSql,dbget,1

%>
<script language="javascript">
function OpenVacationListAdmin(part_sn)
{
	var win = window.open("/admin/member/tenbyten/pop_tenbyten_vacation_list_admin.asp?part_sn=" + part_sn,"OpenVacationListAdmin","width=900,height=500,scrollbars=yes");
	win.focus();
}
</script>
<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr height="25">
		    <td style="border-bottom:1px solid #BABABA">
		        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>휴가신청 승인대기</b>
		    </td>
		    <td align="right" style="border-bottom:1px solid #BABABA">
		        <a href="javascript:OpenVacationListAdmin('')">
		        바로가기
		        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                </a>
		    </td>
		</tr>
		<%	If Not(rsget.EOF or rsget.BOF) then %>
		<tr height="25">
		    <td colspan=2>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="2" class="a">
				<tr align="center">
					<td bgcolor="#DCDCDC">부서</td>
					<td bgcolor="#DCDCDC">건수</td>
					<td bgcolor="#DCDCDC">일수</td>
				</tr>
		<%		Do Until rsget.EOF %>
				<tr align="center">
					<td bgcolor="#EFEFEF" align="left"><%=rsget("part_name")%></td>
					<td bgcolor="#EFEFEF"><%=rsget("cntrequestedday")%> 건</td>
					<td bgcolor="#EFEFEF" align=right>
						<%=rsget("totrequestedday")%> 일
						<a href="javascript:OpenVacationListAdmin(<%=rsget("part_sn")%>)"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
					</td>
				</tr>
		<%
				rsget.MoveNext
				Loop
		%>
				<table>
		<%	else %>
		<tr height="35">
		    <td align="center">승인대기가 없습니다.</td>
		</tr>
		<% end if %>
        </table>
    </td>
</tr>
</table>
<%
	rsget.Close
%>
