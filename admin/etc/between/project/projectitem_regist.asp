<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim opjt, pjt_code
pjt_code = Request("pjt_code")

If pjt_code = "" Then
%>
<script language="javascript">
	alert("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�");
	history.back();
</script>
<%	dbget.close()	:	response.End
End If

Dim pjt_name, pjt_gender, pjt_state, pjt_sortType
Dim strG, strSort

strG  = Request("selG")
strSort  = Request("selSort")

SET opjt = new cProject
	opjt.FRectPjt_code = pjt_code
	opjt.getProjectCont()
	pjt_name		= opjt.FItemList(0).FPjt_name
	pjt_gender		= opjt.FItemList(0).FPjt_gender
	pjt_state		= opjt.FItemList(0).FPjt_state
	pjt_sortType	= opjt.FItemList(0).FPjt_sortType
%>

<script language="javascript">
// ����ǰ �߰� �˾�
function addnewItem(){
	var popwin;
	popwin = window.open("/admin/etc/between/project/pop_project_additemlist.asp?pjt_code=<%=pjt_code%>", "popup_item", "width=1500,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

//��ü����
var ichk;
ichk = 1;
	
function jsChkAll(){			
    var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];
	
		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}

//����
function jsDel(sType, iValue){	
	var frm;		
	var sValue;		
	frm = document.fitem;
	sValue = "";
	
	if (sType ==0) {
		if(!frm.chkI) return;
		
		if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked){
			   	if (sValue==""){
					sValue = frm.chkI[i].value;		
			   	}else{
					sValue =sValue+","+frm.chkI[i].value;		
			   	}	
			}
		}	
		}else{
			if(frm.chkI.checked){
				sValue = frm.chkI.value;
			}	
		}
	
		if (sValue == "") {
			alert('���� ��ǰ�� �����ϴ�.');
			return;
		}
		document.frmDel.itemidarr.value = sValue;
	}else{
		document.frmDel.itemidarr.value = iValue;
	}	
	 
	if(confirm("�����Ͻ� ��ǰ�� �����Ͻðڽ��ϱ�?")){		
		document.frmDel.submit();
	}
}
//�׷�˻�
function jsSearchGroup(){
	document.fitem.submit();	
}
//����
function jsChSort(){
	document.fitem.submit();	
}

//�׷��̵�	

function addGroup(){
	var frm, sValue, sGroup;

	frm = document.fitem;
	sValue = "";
	sGroup =frm.eG.options[frm.eG.selectedIndex].value ;
			
	if(!frm.chkI) return;
	if(!sGroup){
	 alert("�̵��� �׷��� �����ϴ�.");
	 return;
	}
	
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked){
			   if (sValue==""){
				sValue = frm.chkI[i].value;		
				}else{
				sValue =sValue+","+frm.chkI[i].value;		
				}
			}
		}	
	}else{
		sValue = frm.chkI.value;
	}
	
	if (sValue == "") {
		alert('���� ��ǰ�� �����ϴ�.');
		return;
	}
	document.frmG.selGroup.value = frm.eG.options[frm.eG.selectedIndex].value;
	document.frmG.itemidarr.value = sValue;
	document.frmG.submit();
}

// ��ǰ ���� �ϰ� ����
function jsSortSize() {
	var frm;
	var sValue, sSort
	frm = document.fitem;
	sValue = "";
	sSort = "";
		
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(!IsDigit(frm.sSort[i].value)){
				alert("���������� ���ڸ� �����մϴ�.");
				frm.sSort[i].focus();
				return;
			}

			if (sValue==""){
				sValue = frm.chkI[i].value;		
			}else{
				sValue =sValue+","+frm.chkI[i].value;		
			}	
			
			// ���ļ���
			if (sSort==""){
				sSort = frm.sSort[i].value;		
			}else{
				sSort =sSort+","+frm.sSort[i].value;		
			}
		}
	}else{
		sValue = frm.chkI.value;
		if(!IsDigit(frm.sSort.value)){
			alert("���������� ���ڸ� �����մϴ�.");
			frm.sSort.focus();
			return;
		}
		sSort =  frm.sSort.value; 
	}
	document.frmSortSize.itemidarr.value = sValue;
	document.frmSortSize.sortarr.value = sSort;
	document.frmSortSize.submit();
}
function goPage(pg) {
    document.fitem.page.value = pg;
    document.fitem.submit();
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȹ���ڵ�</td>
			<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= pjt_code %></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȹ����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= pjt_name %></td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= getDBcodeByName(opjt.FItemList(0).FPjt_kind) %></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= getDBcodeByName(opjt.FItemList(0).FPjt_state) %></td>
		</tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<% 
					Select Case pjt_gender
						Case "A"	response.write "��ü"
						Case "M"	response.write "����"
						Case "F"	response.write "����"
					End Select
				%>
			</td>
			<td colspan="2" bgcolor="#FFFFFF"></td>
		</tr>
		</table>
	</td>
</tr>
<%
SET opjt = nothing

Dim cPjtGroup, i
SET cPjtGroup = new cProject
	cPjtGroup.FRectPjt_code = pjt_code
	cPjtGroup.getProjectItemGroup()
%>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a">
		<form name="fitem" method="post" action="projectitem_regist.asp">
		<input type="hidden" name="page" value="">
		<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="selGroup" value="">
		<tr align="center"  >
			<td align="left">
	        	 �׷�˻�
	        	<select name="selG" onChange="jsSearchGroup();">
	        	<option value="">��ü</option>
	       	<% If cPjtGroup.FResultCount > 0 Then %>
	       		<option value="0"  <%IF Cstr(strG) = "0" THEN%>selected<%END IF%>>������</option>
	       	<%
	       		For i = 0 to cPjtGroup.FResultCount - 1
	       	%>
	       		<option value="<%=cPjtGroup.FItemList(i).FPjtgroup_code%>" <%IF Cstr(strG) = Cstr(cPjtGroup.FItemList(i).FPjtgroup_code) THEN %> selected<%END IF%>> <%=cPjtGroup.FItemList(i).FPjtgroup_code%>(<%=cPjtGroup.FItemList(i).FPjtgroup_desc%>)</option>
	    	<%	Next
	    	END IF%>
	       	</select>
	        </td>
	        <td align="right">
	         ���� : <select name="selSort" onchange="jsChSort();">
	       		<option value="sitemid" >�Ż�ǰ��</option>
	       		<option value="sevtitem" <%IF Cstr(strSort) = "sevtitem" THEN %>selected<%END IF%>>������</option>
	       		<option value="sbest" <%IF Cstr(strSort) = "sbest" THEN %>selected<%END IF%>>����Ʈ������</option>
	       		<option value="shsell" <%IF Cstr(strSort) = "shsell" THEN %>selected<%END IF%>>�������ݼ�</option>
	       		<option value="slsell" <%IF Cstr(strSort) = "slsell" THEN %>selected<%END IF%>>�������ݼ�</option>
	       		<option value="sevtgroup" <%IF Cstr(strSort) = "sevtgroup" THEN %>selected<%END IF%>>�׷��</option>
	       		<option value="sbrand" <%IF Cstr(strSort) = "sbrand" THEN %>selected<%END IF%>>�귣��</option>
	       		</select>
	        </td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
		<tr height="35">
			<td align="left">
				<input type="button" value="���û���" onClick="jsDel(0,'');" class="button">&nbsp;&nbsp;&nbsp;
				<select name="eG">
		<%
			If cPjtGroup.FResultCount > 0 Then
				For i = 0 to cPjtGroup.FResultCount - 1
		%>
					<option value=" <%=cPjtGroup.FItemList(i).FPjtgroup_code%>" ><%IF cPjtGroup.FItemList(i).FPjtgroup_pcode <> 0 THEN%>��&nbsp;<%END IF%><%=cPjtGroup.FItemList(i).FPjtgroup_code%>(<%=cPjtGroup.FItemList(i).FPjtgroup_desc%>)</option>
		<%
				Next
			ELSE
		%>
					<option value=""> --�׷����--</option>
		<% END IF %>
				</select>
				<input type="button" value="���ñ׷��̵�" onClick="addGroup();" class="button">
			</td>
			<td align="right">
				<input type="button" value="���� ����" onClick="jsSortSize();" class="button">&nbsp;
				<input type="button" value="����ǰ �߰�" onclick="addnewItem();" class="button">
			</td>
		</tr>
		</table>
	</td>
</tr>
<%
SET cPjtGroup = nothing

Dim cPjtGroupItem, page
page    = request("page")
If page = "" Then page = 1

SET cPjtGroupItem = new cProject
	cPjtGroupItem.FPageSize 	= 20
	cPjtGroupItem.FCurrPage		= page
	cPjtGroupItem.FRectPjt_code = pjt_code
	cPjtGroupItem.FRectSGroup 	= strG
	cPjtGroupItem.FRectSort		= strSort
	cPjtGroupItem.getProjectItem()
%>
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="16" align="left">�˻���� : <b><%= FormatNumber(cPjtGroupItem.FTotalCount,0) %></b>&nbsp;&nbsp;������ : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(cPjtGroupItem.FTotalPage,0) %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>
			<td>�׷��ڵ�</td>
			<td align="center">��ǰID</td>
			<td align="center">�̹���</td>
			<td align="center">�귣��</td>
			<td align="center">��ǰ��</td>
			<td align="center">��Ʈ�� ��ǰ��</td>
			<td align="center">�ǸŰ�</td>
			<td align="center">���԰�</td>
			<td align="center">���</td>
			<td align="center">����</td>
			<td align="center">�Ǹſ���</td>
			<td align="center">��뿩��</td>
			<td align="center">��������</td>
			<td>����</td>
			<td>ó��</td>
		</tr>
<%
	If cPjtGroupItem.FResultCount > 0 Then
    	For i = 0 to cPjtGroupItem.FResultCount - 1
%>
		<tr align="center" bgcolor="#FFFFFF">
			<td><input type="checkbox" name="chkI" value="<%= cPjtGroupItem.FItemList(i).FItemid %>"></td>
			<td><%= Chkiif(cPjtGroupItem.FItemList(i).FPjtgroup_code <> 0, cPjtGroupItem.FItemList(i).FPjtgroup_code, "") %></td>
			<td>
				<A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= cPjtGroupItem.FItemList(i).FItemid %>" target="_blank"><%= cPjtGroupItem.FItemList(i).FItemid %></a>
			<% If cPjtGroupItem.IsSoldOut(cPjtGroupItem.FItemList(i).FSellyn, cPjtGroupItem.FItemList(i).FLimityn, cPjtGroupItem.FItemList(i).FLimitno, cPjtGroupItem.FItemList(i).FLimitsold) Then %>
				<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
			<% End If %>
			</td>
	    	<td>
	    		<% If (Not IsNull(cPjtGroupItem.FItemList(i).FSmallimage) ) and (cPjtGroupItem.FItemList(i).FSmallimage <> "") Then %>
					<img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid( cPjtGroupItem.FItemList(i).FItemid )%>/<%=cPjtGroupItem.FItemList(i).FSmallimage%>">
				<% End If %>
			</td>
			<td><%=db2html(cPjtGroupItem.FItemList(i).FMakerid)%></td>
			<td align="left">&nbsp;<%=db2html(cPjtGroupItem.FItemList(i).FItemname)%></td>
			<td align="left">&nbsp;<%=db2html(cPjtGroupItem.FItemList(i).FChgItemname)%></td>
			<td>
			<%
				Response.Write FormatNumber(cPjtGroupItem.FItemList(i).FOrgprice,0)
				'���ΰ�
				If cPjtGroupItem.FItemList(i).FSailyn="Y" then
					Response.Write "<br><font color=#F08050>(��)" & FormatNumber(cPjtGroupItem.FItemList(i).FSailprice,0) & "</font>"
				End If
				'������
				If cPjtGroupItem.FItemList(i).FItemcouponyn = "Y" Then
					Select Case cPjtGroupItem.FItemList(i).FItemcoupontype
						Case "1"
							Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(cPjtGroupItem.FItemList(i).FOrgprice * ((100 - cPjtGroupItem.FItemList(i).FItemcouponvalue) / 100), 0) & "</font>"
						Case "2"
							Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(cPjtGroupItem.FItemList(i).FOrgprice - cPjtGroupItem.FItemList(i).FItemcouponvalue, 0) & "</font>"
					End Select
				End If
			%>
			</td>
	    	<td>
			<%
				Response.Write FormatNumber(cPjtGroupItem.FItemList(i).FOrgsuplycash,0)
				'���ΰ�
				If cPjtGroupItem.FItemList(i).FSailyn = "Y" Then
					Response.Write "<br><font color=#F08050>" & FormatNumber(cPjtGroupItem.FItemList(i).FSailsuplycash,0) & "</font>"
				End If
				'������
				If cPjtGroupItem.FItemList(i).FItemcouponyn = "Y" Then
					If cPjtGroupItem.FItemList(i).FItemcoupontype = "1" OR cPjtGroupItem.FItemList(i).FItemcoupontype = "2" Then
					End If
				End If
			%>
			</td>
	    	<td><%= fnColor(cPjtGroupItem.IsUpcheBeasong(cPjtGroupItem.FItemList(i).FDeliverytype),"delivery")%></td>
	    	<td>
    		<%
				If cPjtGroupItem.FItemList(i).Fsellcash<>0 Then
					response.write CLng(10000-cPjtGroupItem.FItemList(i).Fbuycash/cPjtGroupItem.FItemList(i).Fsellcash*100*100)/100 & "%"
				End If
			%>
	    	</td>
	    	<td><%= fnColor(cPjtGroupItem.FItemList(i).FSellyn, "yn") %></td>
	    	<td><%= fnColor(cPjtGroupItem.FItemList(i).FIsusing, "yn") %></td>
	    	<td><%= fnColor(cPjtGroupItem.FItemList(i).FLimityn, "yn") %></td>
	    	<td><input type="text" name="sSort" value="<%=cPjtGroupItem.FItemList(i).FPjtitem_sort%>" size="4" style="text-align:right;"></td>
	    	<td><input type="button" value="����" onClick="jsDel(1,<%= cPjtGroupItem.FItemList(i).FItemid %>);" class="button"></td>
	    </tr>
<%
	   Next
%>
		<tr height="20">
			<td colspan="17" align="center" bgcolor="#FFFFFF">
			<% If cPjtGroupItem.HasPreScroll Then %>
				<a href="javascript:goPage('<%= cPjtGroupItem.StartScrollPage-1 %>');">[pre]</a>
			<% Else %>
				[pre]
			<% End If %>
			<% For i=0 + cPjtGroupItem.StartScrollPage to cPjtGroupItem.FScrollCount + cPjtGroupItem.StartScrollPage - 1 %>
				<% If i > cPjtGroupItem.FTotalpage Then Exit For %>
				<% If CStr(page) = CStr(i) Then %>
					<font color="red">[<%= i %>]</font>
				<% Else %>
					<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
				<% End If %>
			<% Next %>
			<% If cPjtGroupItem.HasNextScroll Then %>
				<a href="javascript:goPage('<%= i %>');">[next]</a>
			<% Else %>
				[next]
			<% End If %>
			</td>
		</tr>
<%
	   	ELSE
%>
		<tr align="center" bgcolor="#FFFFFF">
			<td height="50" colspan="16">��ϵ� ������ �����ϴ�.</td>
		</tr>
	   <% END IF %>
		</form>
		</table>
	</td>
</tr>
</table>
<% SET cPjtGroupItem = nothing %>
<!-- ���û���--->
<form name="frmDel" method="post" action="projectitem_process.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- �׷��̵�--->
<form name="frmG" method="post" action="projectitem_process.asp">
<input type="hidden" name="mode" value="G">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="selGroup" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- ���� �� �̹���ũ�� ����--->
<form name="frmSortSize" method="post" action="projectitem_process.asp">
<input type="hidden" name="mode" value="S">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="sortarr" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->