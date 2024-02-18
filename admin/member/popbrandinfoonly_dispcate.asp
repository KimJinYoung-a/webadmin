<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	Dim vMakerID, j, vSelectBox, vArrList, opartner, standardCateCode
	vMakerID = requestCheckVar(request("makerid"),100)
	standardCateCode = getUserCStandardcode(vMakerID)
	If isnull(standardCateCode)  Then standardCateCode = "oo"
	vSelectBox = fnStandardDispCateSelectBoxChk(1,"","dispcate1","", standardCateCode)

	SET opartner = New CPartnerUser
		opartner.FRectDesignerID = vMakerID
		vArrList = opartner.fnUserC_GetDispCateList
	SET opartner = Nothing
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(document.frmImg.standardCateCode.value == ""){
			alert('대표 전시카테고리는 반드시 설정하셔야 합니다');
			document.frmImg.standardCateCode.focus();
			return;
		}
		var cnt = document.all.div1.rows.length;
		document.frmImg.filecnt.value = document.all.div1.rows.length;
		document.frmImg.submit();
	}
	
	function AutoInsert() {
		var f = document.all;
	
		var rowLen = f.div1.rows.length;
		var r  = f.div1.insertRow(rowLen++);
		var c0 = r.insertCell(0);
		
		var Html;

		c0.innerHTML = "&nbsp;";
		var inHtml = "<%=Replace(Replace(vSelectBox,chr(34),"'"),vbCrLf,"")%>";
		inHtml = inHtml.replace("dispcate1","dispcate"+rowLen+"")
		c0.innerHTML = inHtml;
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 전시카테고리 설정(1개 이상 설정 가능)</div>
<table width="380" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="/admin/member/popbrandinfoonly_dispcate_proc.asp">
<input type="hidden" name="makerid" value="<%=vMakerID%>">
<input type="hidden" name="filecnt" value="0">
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center" valign="top" style="padding:6 0 0 2">대표전시카테고리</td>
		<td bgcolor="#FFFFFF" colspan="2">
			<%= fnStandardDispCateSelectBox(1,"", "standardCateCode", standardCateCode, "")%>
		</td>
	</td>
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center" valign="top" style="padding:6 0 0 2">전시카테고리</td>
		<td bgcolor="#FFFFFF">
			<table cellpadding="0" cellspacing="0" border="0" id="div1">
			<%
				IF isArray(vArrList) THEN
					For j=0 To UBound(vArrList,2)
						Response.Write "<tr>" & vbCrLf
						Response.Write "	<td>" & vbCrLf
						Response.Write fnStandardDispCateSelectBoxChk(1,"","dispcate"&(j+1),vArrList(0,j), standardCateCode) & vbCrLf
						Response.Write "	</td>" & vbCrLf
						Response.Write "</tr>" & vbCrLf
					Next
				End If
			%>
			<tr>
				<td>
					<%
						IF isArray(vArrList) THEN
							Response.Write Replace(vSelectBox,"dispcate1","dispcate"&j+1)
						Else
							Response.Write vSelectBox
						End If
					%>
				</td>
			</tr>
			</table>
		</td>
		<td bgcolor="#FFFFFF" width="50" align="center" valign="top" style="padding:3 0 0 0">
			<input type="button" value="추가" onClick="AutoInsert()" class="button">
		</td>
	</tr>	
	<tr>
		<td colspan="3" bgcolor="#FFFFFF">
			* 카테고리 삭제시 "-선택-" 을 선택.<span style="width:77px;"></span>
			<!--<input type="image" src="/images/icon_confirm.gif">//-->
			<img src="/images/icon_confirm.gif" style="cursor:pointer" onclick="jsUpload();">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->