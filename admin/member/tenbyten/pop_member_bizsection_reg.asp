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
			 	alert(iTot + "%-���������� ���� 100%�� �ƴմϴ�.�ٽ� �Է����ּ���");
			 	return;
			}
 			document.frm.submit();
	}

	function jsDel() {
		if (confirm("������ �����Ͻðڽ��ϱ�?") != true) {
			return;
		}

		document.frm.hidM.value = "D";
		document.frm.submit();
	}

</script>
<table width="100%" align="left"   cellpadding="5" cellspacing="0" class="a">
<tr>
	<Td>�μ��� �������� ���<br><hr width="100%"> </td>
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
				<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="Center">�̸�</td>
				<td bgcolor="#FFFFFF"><%= username %> (<%= userid %>)</td>
				<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="Center">���</td>
				<td bgcolor="#FFFFFF"><%=sEmpNo%></td>
			</tR>
			<tr>
				<td align="Center"  bgcolor="<%= adminColor("tabletop") %>">��¥</td>
				<td bgcolor="FFFFFF" colspan="3">
							 <%=Year(sDate)%> �� <%=Month(sDate)%> ��
				</td>
			</tr>
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td width="100"  align="Center" rowspan="30">�μ��� ��������</td>
		 		<td colspan="2"  width="320"  align="Center">�μ�</td>
				<td width="160"  align="Center"> ��������</td>
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
									 ��  <input type="hidden" name="sBCD" value="<%=arrList(2,intLoop)%>">
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
		<td>	<font color="red">+ ��� �� ������ �ش� �� <b>10��</b> ���������� �����մϴ�.</font></td>
	</tr-->
	<tr>
		<td align="center" height="50" valign="top">
			<!--%IF day(date()) <= 10 THEN%--><input type="button" value="���" class="button" onClick="jsReg();" style="width:100px"> <!--%END IF%-->
			<% if isArray(arrList) and (delAvail = "Y") and (C_ADMIN_AUTH or (session("ssAdminPsn") = "8")) then %>
				<input type="button" value="����(������)" class="button" onClick="jsDel();" style="width:100px">
			<% end if %>
		</td>
	</tr>
	</table>
	<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
