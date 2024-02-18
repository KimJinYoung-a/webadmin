<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/CategoryMaster/matching2/classes/categoryMatchingCls.asp"-->
<%
'###############################################
' PageName : popMatching.asp
' Discription : ī�װ� ��Ī ���, ����
'###############################################
dim cdl, cdm, cds, dispCate , dispFullName
Dim clsCM,cdl_nm, cdm_nm, cds_nm

dispCate = requestCheckvar(request("disp"),16) 
 
set clsCM = new CCategoryMatching
	clsCM.FRectDispCate = dispCate 
	dispFullName = clsCM.fnGetDispCateFullName	'����ī�װ���
  
	clsCM.fnGetCategoryDisp '��Ī ī�װ�
	cdl = clsCM.FCateLarge
	cdm = clsCM.FCateMid
	cds = clsCM.FCateSmall
	cdl_nm = clsCM.FCateLargeName 
	cdm_nm = clsCM.FCateMidName   
	cds_nm = clsCM.FCateSmallName 

set clsCM = nothing	 
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	function jsChkSubmit(){
		if(!document.frm.cd1.value || !document.frm.cd2.value || !document.frm.cd3.value){
	 	alert("ī�װ��� �������ּ���");
	 	return;
		}  
		
		document.frm.submit();
	}
	
	// ī�װ����
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}


function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.frm;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
 
}
</script>
<table border="0" cellspacing="1" cellpadding="3" width="100%" class="a">
	<tr>
		<td>ī�װ� ��Ī<hr></tD>
	</tr>
	<tr>
		<td> 
			<form name="frm" method="post" action="procMatching.asp">
				<input type="hidden" name="cd1" value="<%=cdl%>">
				<input type="hidden" name="cd2" value="<%=cdm%>">
				<input type="hidden" name="cd3" value="<%=cds%>">
				<input type="hidden" name="disp" value="<%=dispCate%>">
			<table border=0 cellspacing=1 cellpadding=3 width="100%" class=a bgcolor="#808080">
				<tr height="30">
					<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">����ī�װ�</td>
					<td bgcolor="#FFFFFF"><%=replace(dispFullName,"^^"," > ")%></td>
				</tr>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����ī�װ�</td>
					<td bgcolor="#FFFFFF">
						<input type="text" name="cd1_name" value="<%=cdl_nm%>" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
						<input type="text" name="cd2_name" value="<%=cdm_nm%>" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro">
						<input type="text" name="cd3_name" value="<%=cds_nm%>" id="[on,off,off,off][ī�װ�]" size="20" readonly class="text_ro"> 
						<input type="button" value="ī�װ� ����" class="button" onclick="editCategory(frm.cd1.value,frm.cd2.value,frm.cd3.value);">  </td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
	<tr>
		<td  align="center"><input type="button" class="input" value="Ȯ��" onClick="jsChkSubmit();"></td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->