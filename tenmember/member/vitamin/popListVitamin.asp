<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 비타민 등록
' History : 2017.03.14
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVitaminCls.asp" -->
<%
dim clsVM 
dim empno, arrList, intLoop
dim totvm, usevm, reqvm, payvm
empno= session("ssBctSn")

		set clsVM = new CMyVitamin
		clsVM.FRectempno = empno
		clsVM.fnGetMyVitamin
		totvm = clsVM.Ftotvm
		usevm = clsVM.Fusevm 
		clsVM.FRectyyyy = year(date())
		arrList = clsVM.fnGetMyVitaminList
		set clsVM = nothing
%>
<script type="text/javascript">
	 function jsRegEapp(didx, reqvm){
	   var wineapp =window.open("/admin/approval/eapp/regeapp.asp?iAidx=351&ieidx=33&iSL="+didx+"&mRP="+reqvm,"popNew","width=880, height=600,scrollbars=yes, resizable=yes");
	   wineapp.focus();
	 } 
	 
	  function jsViewEapp(iridx){	  	 
	   var winVME =window.open("/admin/approval/eapp/modeapp.asp?iridx="+iridx,"popVM","width=880, height=600,scrollbars=yes, resizable=yes");
	   winVME.focus();
	 }
	 
	 function jsDelVM(didx){
	 	if(confirm("삭제하시겠습니까?")){
	 		document.frmD.didx.value = didx; 
	 		document.frmD.submit();
	 	}
		} 
</script>
<div style="padding:10px;">
비타민 신청내역<br><hr width="100%"> 
</div>
<form name="frmD" method="post" action="procVitamin.asp">
	<input type="hidden" name="didx" value="">
	<input type="hidden" name="hidM" value="D">
</form>
<div style="padding:10px;margin-bottom:10px;"> 
<table width="100%"  cellpadding="10" cellspacing="1" class="a"  bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100">사번</td>
		<td  bgcolor="FFFFFF"><%=empno%></td>
		<td bgcolor="<%= adminColor("tabletop") %>"  width="100">년도</td>
  	<td colspan="3" bgcolor="FFFFFF">금년(<%=year(date())%>)</td>
	</tr>
	<tr> 
  	<td  bgcolor="<%= adminColor("tabletop") %>">부여금액</td>
  	<td bgcolor="FFFFFF"><%=formatnumber(totvm,0)%></td>
  	<td  bgcolor="<%= adminColor("tabletop") %>">사용금액</td> 
  	<td bgcolor="FFFFFF"><%=formatnumber(usevm,0)%></td> 
  	<td  bgcolor="<%= adminColor("tabletop") %>"  width="100">잔액</td>
  	<td bgcolor="FFFFFF"><%=formatnumber(totvm-usevm,0)%></td>
 </tr>
 
</table>	 
</div>
<div>
	<table width="100%"  cellpadding="10" cellspacing="1" class="a"  bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>idx</td>
			<td>신청금액</td>
			<td>신청일</td>
			<td>지급일</td>
			<td>상태</td>
		</tr>
		<% if isArray(arrList) then
				for intLoop = 0 To UBound(arrList,2)				
			%>
			<tr  align="center" bgcolor="FFFFFF">
			<td><%=arrList(0,intLoop)%></td>
			<td><%=formatnumber(arrList(1,intLoop),0)%></td>
			<td><%=arrList(2,intLoop)%></td>
			<td><%=arrList(3,intLoop)%></td>
			<td><%=fnMyStatusDesc(arrList(4,intLoop),arrList(5,intLoop),arrList(0,intLoop),arrList(1,intLoop))%></td>
		</tr>
		<%	next
		end if %>
</div>
 
		
		