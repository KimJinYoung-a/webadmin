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
<!-- #include virtual="/lib/classes/linkedERP/bizsectionCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%

Dim clsBiz, clsMem, cMember
Dim sBizsection_cd
Dim intY, intM, sYear, sMonth
Dim arrList, intLoop
Dim sDate

sBizsection_cd = requestCheckVar(Request("sBcd"),10)
sDate  = requestCheckVar(Request("sD"),7)

	Set clsBiz = new CBizSection
		clsBiz.Fyyyymm = sDate
		clsBiz.FBizsection_cd = sBizsection_cd
		arrList = clsBiz.fnGetManualBizList
	Set clsBiz = nothing

%>
<script type="text/javascript">

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
</script>
<table width="100%" align="left"   cellpadding="5" cellspacing="0" class="a">
<tr>
	<Td>�μ��� �������� ���<br><hr width="100%"> </td>
</tr>
<tr>
	<td>
		<form name="frm" method="post" action="member_bizsection_proc.asp">
		<input type="hidden" name="hidM" value="M">
		<input type="hidden" name="hidUBCD" value="<%=sBizsection_cd%>">
		<table width="580" cellpadding="5" cellspacing="1" class="a" bgcolor="#BABABA">
			<tr>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="Center">�μ�</td>
				<td bgcolor="#FFFFFF" colspan="3">����</td>
			</tr>
			<tr>
				<td align="Center"  bgcolor="<%= adminColor("tabletop") %>">��¥</td>
				<td bgcolor="FFFFFF" colspan="3">
					<select name="selY" class="select">
						<%For intY = Year(date()) To 2011 STEP -1%>
						<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(Year(sDate)) THEN%>selected<%END IF%>><%=intY%></option>
						<%Next%>
						</select>��
						 <select name="selM" class="select">
						<%For intM = 1 To 12%>
						<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(Month(sDate)) THEN%>selected<%END IF%>><%=intM%></option>
						<%Next%>
						</select>��
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
		<td align="center" height="50" valign="top"><!--%IF day(date()) <= 10 THEN%--><input type="button" value="���" class="button" onClick="jsReg();" style="width:100px"> <!--%END IF%--></td>
	</tr>
	</table>
	<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
