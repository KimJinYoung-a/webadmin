<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 운영비관리 - 부서 선택
' History : 2011.06.02 정윤정  생성
'			2018.10.11 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
dim i, j, k
dim oCTenByTenDepartment
set oCTenByTenDepartment = new CTenByTenDepartment
	oCTenByTenDepartment.FPageSize = 500
	oCTenByTenDepartment.FCurrPage = 1
	oCTenByTenDepartment.FRectUseYN = "Y"
	oCTenByTenDepartment.GetList

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<script type="text/javascript">
<!--

function jsAddDepartment() {
 	var oldValue=0;
	var chkValue="";
	var strValue="";
	var arrValue = "";

	if(opener.document.frm.hidDPid.value!="") {
	 	arrValue = opener.document.frm.hidDPid.value.split(",");
	 	oldValue = arrValue.length;
	}

	if(typeof(document.frmReg.chkdpid)=="undefined") {
	 	return;
	}

	if(typeof(document.frmReg.chkdpid.length)=="undefined") {
	  	if(document.frmReg.chkdpid.checked) {
		  	if(oldValue == 1) {
		  		if(arrValue==document.frmReg.chkdpid.value) {
		  			return;
		  		}
		  	}else if(oldValue > 1) {
		  		for(j=0;j<oldValue;j++) {
		  			if(arrValue[j]==document.frmReg.chkdpid.value) {
		  				return;
		  			}
		  		}
		  	}
	  		chkValue = document.frmReg.chkdpid.value;
	  		strValue ="<div id='dDP"+document.frmReg.chkdpid.value+"'><input type='hidden' name='hidDPid' value='"+document.frmReg.chkdpid.value+"'>"+document.frmReg.hidDPName.value+" <a href='javascript:jsDelDepartment("+document.frmReg.chkdpid.value+")'>[X]</a></div>";
	  	}
	}

	for(i=0;i<document.frmReg.chkdpid.length;i++) {
		var chkReturn=0;
		if(document.frmReg.chkdpid[i].checked) {
	  		if(oldValue == 1) {
	  			if(arrValue==document.frmReg.chkdpid[i].value) {
	  				chkReturn = 1;
	  			}
	  		}else if(oldValue > 1) {
	  			for(j=0;j<oldValue;j++) {
	  				if(arrValue[j]==document.frmReg.chkdpid[i].value) {
	  					chkReturn = 1;
	  				}
	  			}
	  		}
	  		if(chkReturn==0) {
				if(chkValue=="") {
					chkValue = document.frmReg.chkdpid[i].value;
		   			strValue = "<div id='dDP"+document.frmReg.chkdpid[i].value+"'>"+document.frmReg.hidDPName[i].value+" <a href='javascript:jsDelDepartment("+document.frmReg.chkdpid[i].value+")'>[X]</a></div>"
				}else{
					chkValue = chkValue +","+ document.frmReg.chkdpid[i].value;
		  			strValue = strValue + "<div id='dDP"+document.frmReg.chkdpid[i].value+"'>"+document.frmReg.hidDPName[i].value+" <a href='javascript:jsDelDepartment("+document.frmReg.chkdpid[i].value+")'>[X]</a></div>";
		  		}
			}
	  	}
	}

	if(chkValue=="") {
	 	alert("추가하실 부서를 선택해주세요");
		return;
	}

	if(opener.document.frm.hidDPid.value =="") {
		opener.document.frm.hidDPid.value = chkValue;
	}else{
		opener.document.frm.hidDPid.value = opener.document.frm.hidDPid.value +","+chkValue;
	}

	opener.document.all.divDepartment.innerHTML = opener.document.all.divDepartment.innerHTML  + strValue;
	self.close();
}

//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
	<td><strong>부서 선택</strong><br><hr width="100%"></td>
</tr>
<tr>
	<tD align="right"><input type="button" value="선택 추가" class="button" onClick="jsAddDepartment();"></td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frmReg" method="post" action="procPart11.asp">
		<tr height="35" bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<td width="50">선택</td>
		 	<td width="50">부서<br>번호</td>
			<td width="400">부서명</td>
		</tr>
		<tr>
			<td colspan="3" align="center"  bgcolor="#FFFFFF" >
				<div style="height:450;overflow:scroll;">
					<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
						<%
						if oCTenByTenDepartment.FResultCount > 0 then
							for i = 0 to oCTenByTenDepartment.FResultcount - 1
								%>
						<tr  bgcolor="#FFFFFF" align="center">
		 					<td width="50"><input type="checkbox" name="chkdpid" value="<%= oCTenByTenDepartment.FItemList(i).Fcid %>"></td>
		 					<td width="50"><%= oCTenByTenDepartment.FItemList(i).Fcid %></td>
		 					<td width="430" align="left">
								&nbsp;
								<%= oCTenByTenDepartment.FItemList(i).FdepartmentNameFull %>
								<input type="hidden" name="hidDPName" value="<%= oCTenByTenDepartment.FItemList(i).FdepartmentNameFull %>">
							</td>
						</tr>
								<%
							next
						end if
						%>
		 			</table>
		 		</div>
		 	</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>
