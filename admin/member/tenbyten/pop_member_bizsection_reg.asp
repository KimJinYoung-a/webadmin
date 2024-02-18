<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :
' History :
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim clsMem, cMember
Dim sEmpNo
Dim intY, intM, sYear, sMonth
Dim arrList, intLoop
Dim sDate
dim userid, username
dim delAvail

sEmpNo = requestCheckVar(Request("sEn"),32)
sDate  = requestCheckVar(Request("sD"),7)
delAvail  = requestCheckVar(Request("delAvail"),7)

Set cMember = new CTenByTenMember
	cMember.Fempno = sEmpNo
	cMember.fnGetMemberData

userid			= cMember.Fuserid
username      	= cMember.Fusername

Set clsMem = new CTenByTenMember
	clsMem.Fempno = sempno
	clsMem.Fyyyymm = sDate
	arrList = clsMem.fnGetUserBizSectionData
Set clsMem = nothing
%>
<script type="text/javascript">
	function jsGetBizcd(){
		var winGB = window.open("/admin/linkedERP/biz/popGetBiz.asp?iType=9","popGB","width=400, height=600, resizable=yes, scrollbars=yes");
		winGB.focus();
	}

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

	function jsDel() {
		if (confirm("정말로 삭제하시겠습니까?") != true) {
			return;
		}

		document.frm.hidM.value = "D";
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
		<input type="hidden" name="sEn" value="<%=sEmpNo%>">
		<input type="hidden" name="selY" value="<%=Year(sDate)%>">
		<input type="hidden" name="selM" value="<%=month(sDate)%>">
		<table width="580" cellpadding="5" cellspacing="1" class="a" bgcolor="#BABABA">
			<tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="Center">이름</td>
				<td bgcolor="#FFFFFF"><%= username %> (<%= userid %>)</td>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="Center">사번</td>
				<td bgcolor="#FFFFFF"><%=sEmpNo%></td>
			</tR>
			<tr>
				<td align="Center"  bgcolor="<%= adminColor("tabletop") %>">날짜</td>
				<td bgcolor="FFFFFF" colspan="3">
							 <%=Year(sDate)%> 년 <%=Month(sDate)%> 월
				</td>
			</tr>
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td width="100"  align="Center" rowspan="30">부서별 업무비율</td>
		 		<td colspan="2"  width="320"  align="Center">부서</td>
				<td width="160"  align="Center"> 업무비율</td>
			</tr>
			<%
				IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
						IF  arrList(4,intLoop) = ""   THEN
						%>
							<tr bgcolor="#FFFFFF"  height=30 >
								<td  colspan="2" ><%=arrList(2,intLoop)%>&nbsp; <%=arrList(3,intLoop)%></td>
								<td></td>
							</tr>
						<%ELSE%>
							<tr height=30 align="center" bgcolor="#FFFFFF">
								<td align="left"  colspan="2" > &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									 └  <input type="hidden" name="sBCD" value="<%=arrList(2,intLoop)%>">
									 <%=arrList(2,intLoop)%>&nbsp; <%=arrList(3,intLoop)%>
									 </td>
									 <td><input type="text" size="3" name="sPR" style="text-align:right;" value="<%IF isnull(arrList(5,intLoop)) then %>0<%else%><%=arrList(5,intLoop)%><%end if%>"> %</td>
							</tr>
				<%		END IF
					Next
				END IF%>
			</table>
			</form>
		</td>
	</tr>
	<!--tr>
		<td>	<font color="red">+ 등록 및 수정은 해당 월 <b>10일</b> 이전까지만 가능합니다.</font></td>
	</tr-->
	<tr>
		<td align="center" height="50" valign="top">
			<!--%IF day(date()) <= 10 THEN%--><input type="button" value="등록" class="button" onClick="jsReg();" style="width:100px"> <!--%END IF%-->
			<% if isArray(arrList) and (delAvail = "Y") and (C_ADMIN_AUTH or (session("ssAdminPsn") = "8")) then %>
				<input type="button" value="삭제(관리자)" class="button" onClick="jsDel();" style="width:100px">
			<% end if %>
		</td>
	</tr>
	</table>
	<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
