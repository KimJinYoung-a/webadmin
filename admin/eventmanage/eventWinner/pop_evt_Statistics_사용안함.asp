<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/pop_evt_Statistics.asp
' Description :  이벤트 응모자 통계
' History : 2007.09.19 김정인
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->
<%
	Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅

	dim evtCode

	evtCode= request("eC")


	dim teenage,twentyA,twentyB,thirty '// 나이별
	dim OCnt,YCnt,Gcnt,BCnt,Vcnt,MCnt,SCnt,FRCnt,FACnt,EtcCnt '// 등급별
	dim MenCnt,WomCnt '// 성별
	dim SeoCnt,GyeCnt,OthCnt '// 지역별

	dim strSQL
	strSQL =" select "&_
			" sum(case "&_
			" 	when datediff(month,n.birthday,getdate())/12 < 20 "&_
			" 		then 1 else 0  end  "&_
			" 	) as teenage /* 19세이하 */ "&_
			" ,sum(case "&_
			" 	when datediff(month,n.birthday,getdate())/12 >= 20 and datediff(month,n.birthday,getdate())/12 < 26 "&_
			" 		then 1 else 0  end  "&_
			" 	) as twentyA /* 20~26 */ "&_
			" ,sum(case "&_
			" 	when datediff(month,n.birthday,getdate())/12 >= 26 and datediff(month,n.birthday,getdate())/12 < 30 "&_
			" 		then 1 else 0  end  "&_
			" 	) as twentyB /* 26~29 */ "&_
			" ,sum(case "&_
			" 	when datediff(month,n.birthday,getdate())/12 >=30 "&_
			" 		then 1 else 0  end  "&_
			" 	) as thirty /* 30세 이상 */ "&_
			" ,sum(case userlevel "&_
			" 	when 5 then 1 else 0  "&_
			" end) as OCnt /* 오렌지 */   "&_
			" ,sum(case userlevel "&_
			" 	when 0   then 1 else 0  "&_
			" end) as YCnt /* 엘로우 */  "&_
			" ,sum(case userlevel "&_
			" 	when 1  then 1 else 0  "&_
			" end) as Gcnt /* 그린 */  "&_
			" ,sum(case userlevel "&_
			" 	when 2  then 1 else 0  "&_
			" end) as BCnt /* 블루 */  "&_
			" ,sum(case userlevel "&_
			" 	when 3  then 1 else 0  "&_
			" end) as Vcnt /* VIP */  "&_
			" ,sum(case userlevel "&_
			" 	when 9  then 1 else 0  "&_
			" end) as MCnt /* Mania */ "&_
			" ,sum(case userlevel "&_
			" 	when 7  then 1 else 0  "&_
			" end) as SCnt /* staff */  "&_
			" ,sum(case userlevel "&_
			" 	when 6  then 1 else 0  "&_
			" end) as FRCnt /* friends */  "&_
			" ,sum(case userlevel "&_
			" 	when 8  then 1 else 0  "&_
			" end) as FACnt /* family */  "&_
			" ,sum(case  "&_
			" 	when userlevel<>5 and userlevel<>0 and userlevel<>1 and userlevel<>2 and userlevel<>3 and userlevel<>9 and userlevel<>7 and userlevel<>6 and userlevel<>8 "&_
			" 	then 1 else 0  "&_
			" end ) as EtcCnt /* 그외등급 */ "&_
			" ,sum(case n.sexflag  "&_
			" 		when 1 then 1 else 0  end "&_
			" 	) as MenCnt /* 남성회원 */ "&_
			" ,sum(case n.sexflag  "&_
			" 		when 2 then 1 else 0  end "&_
			" 	) as WomCnt /* 여성회원 */ "&_
			" ,sum(case z.addr050_si "&_
			" 		when '서울' then 1 else 0   "&_
			" 	end ) as SeoCnt  /* 서울사라미에요 */ "&_
			" ,sum(case z.addr050_si "&_
			" 		when '경기' then 1 else 0   "&_
			" 	end ) as gyeCnt /* '경기사라미에요' */ "&_
			" ,sum(case  "&_
			" 		when z.addr050_si<>'서울' and z.addr050_si<>'경기' then 1 else 0   "&_
			" 	end ) as othCnt /* 지방사라미에요 */ "&_
			" from db_event.dbo.tbl_event_common_comment c "&_
			" join db_user.[dbo].tbl_user_n n "&_
			" 	on c.userid= n.userid "&_
			" join db_user.[dbo].tbl_logindata g "&_
			" 	on g.userid= n.userid "&_
			" join db_zipcode.[dbo].ADDR050TL z "&_
			" 	on n.zipcode = z.addr050_zip1+'-'+z.addr050_zip2 "&_
			" where c.evt_code='" & CStr(evtCode) & "' "


	'response.write strSQL
	rsget.open strSQL ,dbget,1

	if not rsget.eof then
		teenage = Cint(rsget("teenage"))
		twentyA =Cint(rsget("twentyA"))
		twentyB = Cint(rsget("twentyB"))
		thirty = Cint(rsget("thirty"))

		OCnt = Cint(rsget("OCnt"))
		YCnt = Cint(rsget("YCnt"))
		Gcnt = Cint(rsget("Gcnt"))
		BCnt = Cint(rsget("BCnt"))
		Vcnt = Cint(rsget("Vcnt"))
		MCnt = Cint(rsget("MCnt"))
		SCnt = Cint(rsget("SCnt"))
		FRCnt = Cint(rsget("FRCnt"))
		FACnt = Cint(rsget("FACnt"))
		EtcCnt = Cint(rsget("EtcCnt"))

		MenCnt = Cint(rsget("MenCnt"))
		WomCnt = Cint(rsget("WomCnt"))

		SeoCnt = Cint(rsget("SeoCnt"))
		GyeCnt = Cint(rsget("GyeCnt"))
		OthCnt = Cint(rsget("othCnt"))
	end if

	rsget.close

%>
<!-- 표 중간바 시작-->
<Style>
.barright{border-right:1px solid #CCCCCC;}
</style>
<script language="javascript">
function ShowToolBox(iVal)
{
	var mx = document.body.scrollLeft + event.clientX;
	var my = document.body.scrollTop + event.clientY -30;

	var iTool = document.getElementById("tool");
	iTool.innerHTML = iVal;
	iTool.style.left=mx;
	iTool.style.top=my;
	iTool.style.display="";
}

function HideToolBox(){
	var iTool = document.getElementById("tool");
	iTool.style.display="none";
}
</script>
<table width="360" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100" align="center">총 코멘트수</td>
		<td bgcolor="#FFFFFF"><%= MenCnt + WomCnt %>명</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100" align="center">등급별</td>
		<td bgcolor="#FFFFFF">

			<table width="260" border="0" cellpadding="0" cellspacing="0" class="a" rules="rows">
			<% dim totCnt
			totCnt = OCnt+YCnt+Gcnt+BCnt+Vcnt+MCnt+SCnt+FRCnt+FACnt %>
			<% if OCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((OCnt/totCnt)*100) %>%(<%=OCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('ORANGE');" onmouseout="HideToolBox();"><hr color="#F6931B" size="3" width="<%= (OCnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if YCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((YCnt/totCnt)*100) %>%(<%=YCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('YELLOW');" onmouseout="HideToolBox();"><hr color="#FFAE00" size="3" width="<%= (YCnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if Gcnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((Gcnt/totCnt)*100) %>%(<%=Gcnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('GREEN');" onmouseout="HideToolBox();"><hr color="#17C400" size="3" width="<%= (Gcnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if BCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((BCnt/totCnt)*100) %>%(<%=BCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('BLUE');" onmouseout="HideToolBox();"><hr color="#0048FF" size="3" width="<%= (BCnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if Vcnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((Vcnt/totCnt)*100) %>%(<%=Vcnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('VIP');" onmouseout="HideToolBox();"><hr color="#FF0173" size="3" width="<%= (Vcnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if MCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((MCnt/totCnt)*100) %>%(<%=MCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('MANiA');" onmouseout="HideToolBox();"><hr color="#FF0173" size="3" width="<%= (MCnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if SCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((SCnt/totCnt)*100) %>%(<%=SCnt%>)/td>
				<td align="left" onmouseover="ShowToolBox('STAFF');" onmouseout="HideToolBox();"><hr color="#FF0173" size="3" width="<%= (SCnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if FRCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((FRCnt/totCnt)*100) %>%(<%=FRCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('FRIEND');" onmouseout="HideToolBox();"><hr color="#FF0173" size="3" width="<%= (FRCnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if FACnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((FACnt/totCnt)*100) %>%(<%=FACnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('FAMILY');" onmouseout="HideToolBox();"><hr color="#FF0173" size="3" width="<%= (FACnt/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			</table>
		</td>
	</tr>
    <tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100" align="center">연령별</td>
		<td bgcolor="#FFFFFF">
			<table width="260" border="0" cellpadding="0" cellspacing="0" class="a">
			<%
			totCnt = teenage + twentyA + twentyB + thirty %>
			<% if teenage<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((teenage/totCnt)*100) %>%(<%=teenage%>)</td>
				<td align="left" onmouseover="ShowToolBox('19세이하');" onmouseout="HideToolBox();"><hr color="#F6931B" size="3" width="<%= (teenage/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if twentyA<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((twentyA/totCnt)*100) %>%(<%=twentyA%>)</td>
				<td align="left" onmouseover="ShowToolBox('20-26세');" onmouseout="HideToolBox();"><hr color="#FFAE00" size="3" width="<%= (twentyA/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if twentyB<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((twentyB/totCnt)*100) %>%(<%=twentyB%>)</td>
				<td align="left" onmouseover="ShowToolBox('26-29세');" onmouseout="HideToolBox();"><hr color="#17C400" size="3" width="<%= (twentyB/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			<% if thirty<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((thirty/totCnt)*100) %>%(<%=thirty%>)</td>
				<td align="left" onmouseover="ShowToolBox('30세이상');" onmouseout="HideToolBox();"><hr color="#0048ff" size="3" width="<%= (thirty/totCnt)*200 %>"></td>
			</tr>
			<% end if %>
			</table>
		</td>
	</tr>
    <tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100" align="center">지역별</td>
		<td bgcolor="#FFFFFF">
			<table width="260" border="0" cellpadding="0" cellspacing="0" class="a">
			<%
			totCnt = SeoCnt + GyeCnt + OthCnt %>
			<% if SeoCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((SeoCnt/totCnt)*100) %>%(<%=SeoCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('서울');" onmouseout="HideToolBox();"><hr color="#F6931B" size="3" width="<%= Cint((SeoCnt/totCnt)*200) %>"></td>
			</tr>
			<% end if %>
			<% if GyeCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((GyeCnt/totCnt)*100) %>%(<%=GyeCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('경기');" onmouseout="HideToolBox();"><hr color="#FFAE00" size="3" width="<%= Cint((GyeCnt/totCnt)*200) %>"></td>
			</tr>
			<% end if %>
			<% if OthCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((OthCnt/totCnt)*100) %>%(<%=OthCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('지방');" onmouseout="HideToolBox();"><hr color="#17C400" size="3" width="<%= Cint((OthCnt/totCnt)*200) %>"></td>
			</tr>
			<% end if %>
			</table>
		</td>
	</tr>
    <tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100" align="center">성별</td>
		<td bgcolor="#FFFFFF">
			<table width="260" border="0" cellpadding="0" cellspacing="0" class="a">
			<%
			totCnt = MenCnt + WomCnt %>
			<% if MenCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright" class="barright"><%= Cint((MenCnt/totCnt)*100) %>%(<%=MenCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('남성회원');" onmouseout="HideToolBox();"><hr color="#0048FF" size="3" width="<%= Cint((MenCnt/totCnt)*200) %>"></td>
			</tr>
			<% end if %>
			<% if WomCnt<>"0" then %>
			<tr>
				<td align="left" width="50" class="barright"><%= Cint((WomCnt/totCnt)*100) %>%(<%=WomCnt%>)</td>
				<td align="left" onmouseover="ShowToolBox('여성회원');" onmouseout="HideToolBox();"><hr color="#FF0173" size="3" width="<%= Cint(((WomCnt)/totCnt)*200) %>"></td>
			</tr>
			<% end if %>
			</table>
		</td>
	</tr>

</table>
<div id="tool" style="POSITION: absolute;"></div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->