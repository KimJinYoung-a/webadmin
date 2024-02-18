<%

dim IsUpdateSCMMemberBirthDayInfoNeeded

IsUpdateSCMMemberBirthDayInfoNeeded = False

If Trim(application("scmTimeMemberBirthDayInfo")) = "" Or DateDiff("h", application("scmTimeMemberBirthDayInfo"), Now() ) > 4 Then
	'// 4시간에 한번
	application("scmTimeMemberBirthDayInfo") = Now()
	IsUpdateSCMMemberBirthDayInfoNeeded = True
end if

 
'// 직원 생일자 목록 (금주 해당 표시) //u.posit_sn<13 월급계약직 이상으로 변경
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

	' 퇴사예정자 처리	' 2018.10.16 한용민
	brthSql = brthSql & " 		and (u.statediv ='Y' or (u.statediv ='N' and datediff(dd,u.retireday,getdate())<=0))" & vbcrlf
	brthSql = brthSql & "		and u.part_sn not in (2,3,17)" & vbcrlf		' 공급자 및 관계사 제외
	brthSql = brthSql & "		and u.userid not in ('logicsmulti','rpabot1','rpabot2','10x10staff','iiitester1','iiitester2') " & vbcrlf 		' 봇/테스터 제외
	brthSql = brthSql & " 	) USR "
	brthSql = brthSql & " where "

	'/1주였다가 2주로 바꿈. 자꾸 자기 생일안나온다고 물어봄.	'/2017.07.19 용만
	brthSql = brthSql & " 	datediff(week,'2008-'+right(convert(varchar(10),birthDt,21),5),'2008-'+right(convert(varchar(10),getdate(),21),5)) between -2 and 0 "
	brthSql = brthSql & " order by right(convert(varchar(10),birthDt,21),5)"

	'response.write brthSql & "<br>"
	rsget.Open brthSql,dbget,1

	If Not(rsget.EOF or rsget.BOF) then
		Do Until rsget.EOF
			''직급제거
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
		        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>축하합니다!</b> - 금주 및 차주의 생일자
		    </td>
		</tr>
		<!-- 첫번째 값은 빈값이므로 스킵한다. -->
		<%	If UBound(scmArrMemberBirthDayInfo) >= 1 then %>
		<tr height="25">
		    <td>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="2" class="a">
				<tr align="center">
					<td bgcolor="#DCDCDC">생일</td>
					<td bgcolor="#DCDCDC">부서</td>
					<td bgcolor="#DCDCDC">이름</td>
				</tr>
			<% for i = 1 to UBound(scmArrMemberBirthDayInfo) %>
				<%
				scmItemMemberBirthDayInfo = scmArrMemberBirthDayInfo(i)
				scmItemMemberBirthDayInfo = Split(scmItemMemberBirthDayInfo, "|")
				%>
				<tr align="center">
					<td bgcolor="#EFEFEF">
					<%
						'직급제거로 번호 이동
						'Response.Write Month(scmItemMemberBirthDayInfo(4)) & "월 " & day(scmItemMemberBirthDayInfo(4)) & "일 "
						'if scmItemMemberBirthDayInfo(6)="N" then
						Response.Write Month(scmItemMemberBirthDayInfo(2)) & "월 " & day(scmItemMemberBirthDayInfo(2)) & "일 "
						if scmItemMemberBirthDayInfo(4)="N" then
							Response.Write "[음]"
						else
							Response.Write "[양]"
						end if
					%>
					</td>
					<td bgcolor="#EFEFEF" align="left"><%= replace(scmItemMemberBirthDayInfo(0),"텐바이텐 - ","") %></td>
					<td bgcolor="#EFEFEF">
					<%
						'직급제거
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
		    <td align="center">생일자가 없습니다.</td>
		</tr>
		<% end if %>
        </table>
    </td>
</tr>
</table>
