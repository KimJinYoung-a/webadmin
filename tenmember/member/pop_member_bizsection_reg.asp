<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :
' History :
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%

dim empno
empno = session("ssBctSn")

Dim clsMem
Dim intY, intM, sYear, sMonth
Dim arrList, intLoop
Dim sDate
 sDate  = requestCheckVar(Request("sD"),7)

Set clsMem = new CTenByTenMember
	clsMem.Fempno = empno
	clsMem.Fyyyymm = sDate
	arrList = clsMem.fnGetUserBizSectionData
Set clsMem = nothing
%>
<script type="text/javascript">

function jsReg(){
	var iTot;
	iTot = 0
	for(i=0;i<document.frm.sPR.length;i++){
		iTot = iTot + parseInt(document.frm.sPR[i].value,10);
	}

	if(iTot !=  100){
		alert(iTot + "%-업무비율의 합이 100%가 아닙니다.다시 입력해주세요");
		return;
	}

	document.frm.submit();
}

</script>
<table width="100%" align="left"   cellpadding="5" cellspacing="0" class="a">
<tr>
	<Td>부서별 업무비율 등록<br><hr width="100%"> </td>
</tr>
<tr>
	<td>
		<form name="frm" method="post" action="member_bizsection_proc.asp">
		<input type="hidden" name="hidM" value="I">
		<input type="hidden" name="selY" value="<%=Year(sDate)%>">
		<input type="hidden" name="selM" value="<%=month(sDate)%>">
		<table width="580" cellpadding="5" cellspacing="1" class="a" bgcolor="#BABABA">
			<tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="Center">이름</td>
				<td bgcolor="#FFFFFF"><%=session("ssBctCname")%> (<%=session("ssBctId")%>)</td>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="Center">사번</td>
				<td bgcolor="#FFFFFF"><%= empno %></td>
			</tR>
			<tr>
				<td align="Center"  bgcolor="<%= adminColor("tabletop") %>">날짜</td>
				<td bgcolor="FFFFFF" colspan="3">
							 <%=Year(sDate)%> 년 <%=Month(sDate)%> 월
				</td>
			</tr>
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td width="100"  align="Center" rowspan="2">부서별 업무비율</td>
		 		<td colspan="2"  width="320"  align="Center">부서</td>
				<td    align="Center"> 업무비율</td>
			</tr>
			<tr>
				<td colspan="3" bgcolor="#ffffff">
					<table  width="460" cellpadding="5" cellspacing="1" class="a" bgcolor="#BABABA">
			<%
				IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
						IF  arrList(4,intLoop) = ""   THEN
					%>
					<tr bgcolor="#FFFFFF"  height=30 >
						<td width="310"><%=arrList(6,intLoop)%>&nbsp; <%=arrList(3,intLoop)%></td>
						<td></td>
					</tr>
					<%ELSE%>
					<tr height=30 align="center" bgcolor="#FFFFFF">
						<td align="left"   > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						 └  <input type="hidden" name="sBCD" value="<%=arrList(2,intLoop)%>">
						 <input type="hidden" name="sBUCD" value="<%=arrList(6,intLoop)%>">
						 <%=arrList(6,intLoop)%>&nbsp; <%=arrList(3,intLoop)%>
						 </td>
						 <td><input type="text" size="3" name="sPR" style="text-align:right;" value="<%IF isnull(arrList(5,intLoop)) then %>0<%else%><%=arrList(5,intLoop)%><%end if%>"> %</td>
					</tr>
			<%		END IF
					Next
			END IF%>
					</table>
					</td>
					</tr>
			</table>
			</form>
		</td>
	</tr>
	<tr>
		<td align="center" height="50" valign="top"><input type="button" value="등록" class="button" onClick="jsReg();" style="width:100px"></td>
	</tr>
	</table>
	<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->