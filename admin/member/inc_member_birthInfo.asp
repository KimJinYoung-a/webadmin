<%

dim IsUpdateSCMMemberBirthDayInfoNeeded

IsUpdateSCMMemberBirthDayInfoNeeded = False

If Trim(application("scmTimeMemberBirthDayInfo")) = "" Or DateDiff("h", application("scmTimeMemberBirthDayInfo"), Now() ) > 4 Then
	'// 4�ð��� �ѹ�
	application("scmTimeMemberBirthDayInfo") = Now()
	IsUpdateSCMMemberBirthDayInfoNeeded = True
end if

 
'// ���� ������ ��� (���� �ش� ǥ��) //u.posit_sn<13 ���ް���� �̻����� ����
Dim brthSql, scmArrMemberBirthDayInfo, scmItemMemberBirthDayInfo

if (IsUpdateSCMMemberBirthDayInfoNeeded = True) then
	brthSql = "select * "
	brthSql = brthSql & " from ( "
	brthSql = brthSql & "		select  "

	''brthSql = brthSql & "		t2.part_name,  "
	brthSql = brthSql & "		isNull(dv.departmentNameFull,'') as part_name, "

	brthSql = brthSql & "		u.posit_sn,  "
	brthSql = brthSql & "		t3.posit_name,  "
	brthSql = brthSql & "		u.username as company_name , "
'	brthSql = brthSql & "		Case u.issolar  "
'	brthSql = brthSql & "			When 'Y' Then convert(varchar(10),u.birthday,21)  "
'	brthSql = brthSql & "			When 'N' Then (Select top 1 solar_date from db_sitemaster.dbo.LunarToSolar Where lunar_date=convert(varchar(10),u.birthday,21)) Else convert(varchar(10),u.birthday,21)  "
'	brthSql = brthSql & "		End as birthDt,  "
	brthSql = brthSql & "		convert(varchar(10),u.birthday,21) as birthDt, "
	brthSql = brthSql & "		isNull(u.birthday,'') as birthday, "
	brthSql = brthSql & "		u.issolar as birth_isSolar  "
	brthSql = brthSql & "		from db_partner.dbo.tbl_user_tenbyten as u "
	brthSql = brthSql & "		left join  db_partner.dbo.tbl_partner as t1 on u.userid =t1.id and t1.userdiv<999 and t1.isusing='Y' "
	brthSql = brthSql & "		left join db_partner.dbo.tbl_partInfo as t2 on u.part_sn =t2.part_sn  "
	brthSql = brthSql & "		left join db_partner.dbo.tbl_positInfo as t3 on u.posit_sn=t3.posit_sn  "
	brthSql = brthSql & "		left join db_partner.dbo.vw_user_department dv on u.department_id = dv.cid "
	brthSql = brthSql & "		where  u.isusing =1 and u.posit_sn<13 "

	' ��翹���� ó��	' 2018.10.16 �ѿ��
	brthSql = brthSql & " 		and (u.statediv ='Y' or (u.statediv ='N' and datediff(dd,u.retireday,getdate())<=0))" & vbcrlf
	brthSql = brthSql & "		and u.part_sn not in (2,3,17)" & vbcrlf		' ������ �� ����� ����
	brthSql = brthSql & "		and u.userid not in ('logicsmulti','rpabot1','rpabot2','10x10staff','iiitester1','iiitester2') " & vbcrlf 		' ��/�׽��� ����
	brthSql = brthSql & " 	) USR "
	brthSql = brthSql & " where "

	'/1�ֿ��ٰ� 2�ַ� �ٲ�. �ڲ� �ڱ� ���Ͼȳ��´ٰ� ���.	'/2017.07.19 �븸
	brthSql = brthSql & " 	datediff(week,'2008-'+right(convert(varchar(10),birthDt,21),5),'2008-'+right(convert(varchar(10),getdate(),21),5)) between -2 and 0 "
	brthSql = brthSql & " order by right(convert(varchar(10),birthDt,21),5)"

	'response.write brthSql & "<br>"
	rsget.Open brthSql,dbget,1

	If Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			''��������
			''scmItemMemberBirthDayInfo = CStr(rsget("part_name")) + "|" + CStr(rsget("posit_sn")) + "|" + CStr(rsget("posit_name")) + "|" + CStr(rsget("company_name")) + "|" + CStr(rsget("birthDt")) + "|" + CStr(rsget("birthday")) + "|" + CStr(rsget("birth_isSolar"))
			scmItemMemberBirthDayInfo = CStr(rsget("part_name")) + "|" + CStr(rsget("company_name")) + "|" + CStr(rsget("birthDt")) + "|" + CStr(rsget("birthday")) + "|" + CStr(rsget("birth_isSolar"))

			scmArrMemberBirthDayInfo = scmArrMemberBirthDayInfo + "=|=" +scmItemMemberBirthDayInfo
			rsget.MoveNext
		Loop
	end if
	rsget.Close

	application("scmArrMemberBirthDayInfo") = scmArrMemberBirthDayInfo

else

	scmArrMemberBirthDayInfo = application("scmArrMemberBirthDayInfo")

end If

scmArrMemberBirthDayInfo = Split(scmArrMemberBirthDayInfo, "=|=")

%>
<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr height="25">
		    <td style="border-bottom:1px solid #BABABA">
		        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�����մϴ�!</b> - ���� �� ������ ������
		    </td>
		</tr>
		<!-- ù��° ���� ���̹Ƿ� ��ŵ�Ѵ�. -->
		<%	If UBound(scmArrMemberBirthDayInfo) >= 1 then %>
		<tr height="25">
		    <td>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="2" class="a">
				<tr align="center">
					<td bgcolor="#DCDCDC">����</td>
					<td bgcolor="#DCDCDC">�μ�</td>
					<td bgcolor="#DCDCDC">�̸�</td>
				</tr>
			<% for i = 1 to UBound(scmArrMemberBirthDayInfo) %>
				<%
				scmItemMemberBirthDayInfo = scmArrMemberBirthDayInfo(i)
				scmItemMemberBirthDayInfo = Split(scmItemMemberBirthDayInfo, "|")
				%>
				<tr align="center">
					<td bgcolor="#EFEFEF">
					<%
						'�������ŷ� ��ȣ �̵�
						'Response.Write Month(scmItemMemberBirthDayInfo(4)) & "�� " & day(scmItemMemberBirthDayInfo(4)) & "�� "
						'if scmItemMemberBirthDayInfo(6)="N" then
						Response.Write Month(scmItemMemberBirthDayInfo(2)) & "�� " & day(scmItemMemberBirthDayInfo(2)) & "�� "
						if scmItemMemberBirthDayInfo(4)="N" then
							Response.Write "[��]"
						else
							Response.Write "[��]"
						end if
					%>
					</td>
					<td bgcolor="#EFEFEF" align="left"><%= replace(scmItemMemberBirthDayInfo(0),"�ٹ����� - ","") %></td>
					<td bgcolor="#EFEFEF">
					<%
						'��������
						'Response.Write scmItemMemberBirthDayInfo(3)
						'if scmItemMemberBirthDayInfo(1)<12 then
						'	Response.Write " " & scmItemMemberBirthDayInfo(2)
						'end if
						Response.Write scmItemMemberBirthDayInfo(1)
					%>
					</td>
				</tr>
			<% next %>
				</table>
			</td>
		</tr>
		<%	else %>
		<tr height="35">
		    <td align="center">�����ڰ� �����ϴ�.</td>
		</tr>
		<% end if %>
        </table>
    </td>
</tr>
</table>
