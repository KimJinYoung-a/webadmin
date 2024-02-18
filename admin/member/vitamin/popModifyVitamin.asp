<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVitaminCls.asp" -->
<%
dim idx, menupos, clsvm, arrList
idx = requestCheckvar(request("idx"),8)
menupos =requestCheckvar(request("menupos"),8)

set clsvm	= new Cvitamin
	clsvm.FCurrPage 		= 1
	clsvm.FPageSize 		= 1		
	clsvm.FRectIdx = idx
	arrList = clsvm.fnvitaminGetList
set clsvm	= nothing

if not isArray(arrList) then
	Call alert_return("비타민 없음")
	response.End
end if

%>
<script type="text/javascript">

    function comma(str) {
        str = String(str);
        return str.replace(/(\d)(?=(?:\d{3})+(?!\d))/g, '$1,');
    }

    function uncomma(str) {
        str = String(str);
        return str.replace(/[^\d]+/g, '');
    } 
    
    function inputNumberFormat(obj) {
        obj.value = comma(uncomma(obj.value));
    }

	function recalVitamin() {
		var ototVm = document.frm.totvm;
		var oUseVm = document.frm.usevm;
		var oRmnVm = document.frm.remainvm;

		oRmnVm.value = uncomma(ototVm.value)-uncomma(oUseVm.value);
		if(uncomma(oRmnVm.value)<0) {
			oRmnVm.value = 0;
		}
		if(uncomma(ototVm.value)<uncomma(oUseVm.value)) {
			ototVm.value = oUseVm.value;
		}
		inputNumberFormat(ototVm);
		inputNumberFormat(oRmnVm);
	}

	function frmSubmit() {
		if(document.frm.totvm.value==""){
			alert("비타민을 입력해주세요.");
			document.frm.totvm.focus();
			return;
		}
		if(confirm("입력하신 비타민으로 저장하시겠습니까?")){
			document.frm.totvm.value=uncomma(document.frm.totvm.value);
			document.frm.submit();
		}
	}
</script>
<form name="frm" method="POST" action="procVitamin.asp">
<input type="hidden" name="menupos" value="<%= menupos %>" />
<input type="hidden" name="hidM" value="U" />
<input type="hidden" name="idx" value="<%=idx%>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">사번</td>
	<td><%=arrList(1,0)%></td>
	<td width="80" bgcolor="<%= adminColor("gray") %>">이름</td>
	<td><%=arrList(3,0)%></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">부서</td>
	<td colspan="3"><%=arrList(5,0)%></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">사용가능기간</td>
	<td colspan="3"><%=formatdate(arrList(7,0),"0000-00-00")%>~<%=formatdate(arrList(8,0),"0000-00-00")%></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">총비타민</td>
	<td colspan="3"><input type="text" name="totvm" value="<%=formatnumber(arrList(9,0),0)%>" style="text-align:right; width:100px;" class="text" onkeyup="recalVitamin();"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">사용비타민</td>
	<td colspan="3"><input type="text" name="usevm" value="<%=formatnumber(arrList(10,0),0)%>" style="text-align:right; width:100px;" class="text_ro" readonly="readonly"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">잔여비타민</td>
	<td colspan="3"><input type="text" name="remainvm" value="<%=formatnumber(arrList(9,0)-arrList(10,0),0)%>" style="text-align:right; width:100px;" class="text_ro" readonly="readonly"></td>
</tr>
</table>
<div style="text-align:center; margin-top:10px;"><input type="button" class="button" value="저 장" onclick="frmSubmit()" /></div>
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->