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
dim empno
empno= session("ssBctSn")
dim clsVM
		dim totvm, usevm 
		set clsVM = new CMyVitamin
		clsVM.FRectempno = empno
		clsVM.fnGetMyVitamin
		totvm = clsVM.Ftotvm
		usevm = clsVM.Fusevm 
		set clsVM = nothing
%>
<script type="text/javascript">
	function jsSubmit(){
		var reqvm = document.frmReg.reqVM.value;
		if (!reqvm ){
			alert("신청금액을 입력해주세요");
			document.frmReg.reqVM.focus();
			return;
		}
		
		if(isNaN(reqvm)){
		 alert("숫자만 입력가능합니다.");
		 document.frmReg.reqVM.focus();
		 return;
		}
		 
	 if (parseInt(document.frmReg.hidlvm.value) < parseInt(reqvm)){
	 	alert("잔액과 같거나 적은 금액으로 신청해주세요");
	 	document.frmReg.reqVM.focus();
	 	return;
	 }
	 
	 document.frmReg.submit();
	}
</script>
<div style="padding:10px;">
비타민 등록<br><hr width="100%"> 
</div>
<div style="padding:10px;margin-bottom:10px;">
<form name="frmReg" method="post" action="procVitamin.asp">
	<input type="hidden" name="hidM" value="I">
	<input type="hidden" name="hidlvm" value="<%=totvm-usevm%>">
<table width="100%"  cellpadding="10" cellspacing="1" class="a"  bgcolor="<%= adminColor("tablebg") %>">
<tr  height="30">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100"  align="center">사번</td>
	<td bgcolor="#ffffff"><%=empno%></td>
</tr>
<tr height="30">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center" >잔액</td>
	<td bgcolor="#ffffff"><%=formatnumber(totvm-usevm,0)%></td>
</tr> 
<tr  height="30">
	<td bgcolor="<%= adminColor("tabletop") %>"  align="center">신청금액</td>
	<td bgcolor="#ffffff"><input type="text" name="reqVM" size="10" class="input" style="text-align:right;"></td>
</tr>
</table>	
</form> 
</div>
<div style="width:100%; text-align:center;"> 
	<input type="button" class="button" value="등록" onClick="jsSubmit();">
</div>
		