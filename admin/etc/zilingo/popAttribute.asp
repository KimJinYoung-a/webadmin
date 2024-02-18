<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/zilingo/zilingocls.asp"-->
<%
Dim itemid, itemoption, catekey, oZilingo, i, tmpDepth2Name, j, rowDataCnt, tmpGroupName, selectboxYn
Dim arrAttributeGroup, arrAttributeGroup2, isMustFields, getChgItemname
Dim arrCapacitiGroup, regedAttrs, regedstring, mktMappingType
Dim sizeGroup
Dim colorGroup

itemid = request("itemid")
itemoption = request("itemoption")
catekey = request("catekey")


If itemid = "" OR catekey = "" OR itemoption = "" Then
	Alert_Close("카테고리 및 상품 및 옵션코드 값이 없습니다.")
	response.end
End If

Set oZilingo = new CZilingo
	oZilingo.FRectCatekey = catekey
	oZilingo.FRectGubun = "attributeChoices"
	
	oZilingo.FRectItemid		= itemid
	oZilingo.FRectItemoption	= itemoption

	getChgItemname		= oZilingo.fnChgItemname
	arrAttributeGroup	= oZilingo.getAttributeGroupList
	arrAttributeGroup2	= oZilingo.getAttributeGroupList2
	regedAttrs			= oZilingo.getRegedAttributes
	mktMappingType		= oZilingo.getMktMappingType
Set oZilingo = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function removeLastComma(str) {
   return str.replace(/,(\s+)?$/, '');   
}
function frm_check(){
	var inputArray = "";
	var dataArray = $("#frm").serializeArray(),
	len = dataArray.length,
	dataObj = {};

	for (i=0; i<len; i++) {
		if(dataArray[i].value != ""){
			inputArray = inputArray + dataArray[i].value + ','
		}
	}
	inputArray = removeLastComma(inputArray);
	var tmpArray = new Array();
	tmpArray = inputArray.split(',');

	var needArray = new Array();
	var strNeedFields = $("#isMustFields").val();
	needArray = strNeedFields.split(',');
	
	var spTmpVal;
	var oStr = "";

	for (j = 0; j < needArray.length; j++){
		for(k = 0; k < tmpArray.length; k++){
			spTmpVal = tmpArray[k].split('||')[0]
			if (spTmpVal == needArray[j]){
				oStr = oStr + spTmpVal + ','
				break;
			}
		}
	}
	oStr = removeLastComma(oStr);
	var chkArray = new Array();
	if (oStr != ""){
		chkArray = oStr.split(',');
	}
	if (chkArray.length != needArray.length){
		alert('필수값을 입력하세요');
		return;
	}else{
		if(confirm("저장 하시겠습니까?")) {
			document.tfrom.target = "xLink";
	        document.tfrom.attrs.value = inputArray;
	        document.tfrom.action = "/admin/etc/zilingo/procAttrsZilingo.asp"
	        document.tfrom.submit();
		}
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
	<td bgcolor="<%= adminColor("tabletop") %>"><%= getChgItemname %></td>
</tr>
<tr align="center">
	<td bgcolor="<%= adminColor("tabletop") %>">링크</td>
	<td bgcolor="<%= adminColor("tabletop") %>"><a href="http://www.10x10.co.kr/<%=itemid%>" target="_blank">텐바이텐 링크 클릭</a></td>
</tr>
<form name="frm" id="frm">
<!-- arrAttributeGroup -->
<%
If isArray(arrAttributeGroup) Then
	For i = 0 To Ubound(arrAttributeGroup, 2)
		If arrAttributeGroup(2, i) = "N" Then
			isMustFields = isMustFields & arrAttributeGroup(0, i) & ","
		End If

		If i = 0 Then
%>
<tr align="center" >
	<td colspan="2" bgcolor="<%= adminColor("tabletop") %>"><%= arrAttributeGroup(3, i) %></td>
</tr>
<%
		End If
%>
<tr align="center" >
	<td width="20%" bgcolor="<%= adminColor("tabletop") %>"><%= arrAttributeGroup(1, i) %><%= Chkiif(arrAttributeGroup(2, i) = "N", "<font color='RED'>(*)</font>", "") %></td>
	<td width="80%" bgcolor="#FFFFFF" align="left">
<%
		rowDataCnt = 0
		selectboxYn = ""
		For j = 0 to Ubound(arrAttributeGroup2, 2)
			If isArray(arrAttributeGroup2) Then
				If arrAttributeGroup2(5, j) = "Y" Then
					If arrAttributeGroup(0, i) = arrAttributeGroup2(0, j) Then
						rowDataCnt = rowDataCnt + 1
						regedstring= ""
						If Instr(regedAttrs, arrAttributeGroup(0, i)&"||"&arrAttributeGroup2(2, j)) > 0 Then
							regedstring = "checked"
						End If
%>
			<label><input type="checkbox" <%= regedstring %> name="check_<%= arrAttributeGroup(0, i) %>" value="<%= arrAttributeGroup(0, i)&"||"&arrAttributeGroup2(2, j) %>"><%= arrAttributeGroup2(3, j) %></label>&nbsp;
<%
						If rowDataCnt mod 5 = 0 Then response.write "<br />"
					End If
				Else
					If arrAttributeGroup(0, i) = arrAttributeGroup2(0, j) Then
						If rowDataCnt = 0 Then
							response.write "<select name='select_"&arrAttributeGroup(0, i)&"' class='select'><option value=''>-Choice-</option> "
						End If
						rowDataCnt = rowDataCnt + 1

						regedstring= ""

						If Instr(regedAttrs, arrAttributeGroup(0, i)&"||"&arrAttributeGroup2(2, j)) > 0 Then
%>
			<option selected value="<%= arrAttributeGroup(0, i)&"||"&arrAttributeGroup2(2, j) %>"><%= arrAttributeGroup2(3, j) %></option>
<%
							Exit For
						Else
							If Instr(arrAttributeGroup2(2, j), mktMappingType) > 0 Then
%>
			<option selected value="<%= arrAttributeGroup(0, i)&"||"&arrAttributeGroup2(2, j) %>"><%= arrAttributeGroup2(3, j) %></option>
<%
								Exit For
							Else
%>
			<option value="<%= arrAttributeGroup(0, i)&"||"&arrAttributeGroup2(2, j) %>"><%= arrAttributeGroup2(3, j) %></option>
<%
							End If
						End If
%>
<%
						If rowDataCnt mod 5 = 0 Then response.write "<br />"
						selectboxYn = "Y"
					End If
				End If
			End If
		Next

		If selectboxYn = "Y" Then
			 response.write "</select>"
		End If
%>
	</td>
</tr>
<%
	Next
End If

If Right(isMustFields,1) = "," Then
	isMustFields = Left(isMustFields, Len(isMustFields) - 1)
End If
%>
<!-- colorGroup -->
<%
Set oZilingo = new CZilingo
	oZilingo.FRectCatekey = catekey
	oZilingo.FRectGubun = "colors"
	colorGroup = oZilingo.getAttributeGroupList
Set oZilingo = nothing

If isArray(colorGroup) Then
	isMustFields = isMustFields & ",COLORS"
%>
<tr align="center">
	<td colspan="2" bgcolor="RED"></td>
</tr>
<tr align="center" >
	<td width="20%" bgcolor="<%= adminColor("tabletop") %>">colors<font color='RED'>(*)</font></td>
	<td width="80%" bgcolor="#FFFFFF" align="left">
		<select class="select" class="select" name="colors">
			<option value="">-Choice-</option>
	<%
		For i = 0 to Ubound(colorGroup, 2)
			regedstring= ""
			If Instr(regedAttrs, "COLORS||"&colorGroup(0, i)) > 0 Then
				regedstring = "selected"
			End If
	%>
			<option <%= regedstring %> value="COLORS||<%= colorGroup(0, i) %>"><%= colorGroup(1, i) %></option>
	<% Next %>
		</select>
	</td>
</tr>
<% End If %>

<!-- sizeGroup  -->
<%
Set oZilingo = new CZilingo
	oZilingo.FRectCatekey = catekey
	oZilingo.FRectGubun = "sizes"
	sizeGroup = oZilingo.getAttributeGroupList
Set oZilingo = nothing

If isArray(sizeGroup) Then
	isMustFields = isMustFields & ",SIZES"
%>
<tr align="center">
	<td colspan="2" bgcolor="RED"></td>
</tr>
<tr align="center" >
	<td width="20%" bgcolor="<%= adminColor("tabletop") %>">sizes<font color='RED'>(*)</font></td>
	<td width="80%" bgcolor="#FFFFFF" align="left">
		<select class="select" class="select" name="sizes">
			<option value="">-Choice-</option>
	<%
	For i = 0 to Ubound(sizeGroup, 2)
		regedstring= ""
		If Instr(regedAttrs, "SIZES||"&sizeGroup(0, i)) > 0 Then
			regedstring = "selected"
		End If
	%>
			<option <%= regedstring %> value="SIZES||<%= sizeGroup(0, i) %>"><%= sizeGroup(1, i) %></option>
	<% Next %>
		</select>
	</td>
</tr>
<% End If %>

<!-- capacitiGroup  -->
<%
Set oZilingo = new CZilingo
	oZilingo.FRectCatekey = catekey
	oZilingo.FRectGubun = "capacities"
	arrCapacitiGroup = oZilingo.getAttributeGroupList
Set oZilingo = nothing

If isArray(arrCapacitiGroup) Then
	isMustFields = isMustFields & ",CAPACITIES"
%>
<tr align="center">
	<td colspan="2" bgcolor="RED"></td>
</tr>
<tr align="center" >
	<td width="20%" bgcolor="<%= adminColor("tabletop") %>">capacities<font color='RED'>(*)</font></td>
	<td width="80%" bgcolor="#FFFFFF" align="left">
		<select class="select" class="select" name="capacities">
			<option value="">-Choice-</option>
	<%
	For i = 0 to Ubound(arrCapacitiGroup, 2)
		regedstring= ""
		If Instr(regedAttrs, "CAPACITIES||"&arrCapacitiGroup(0, i)) > 0 Then
			regedstring = "selected"
		End If
	%>
			<option <%= regedstring %> value="CAPACITIES||<%= arrCapacitiGroup(0, i) %>"><%= arrCapacitiGroup(1, i) %></option>
	<% Next %>
		</select>
	</td>
</tr>
<% End If %>
</form>
<tr align="center" >
	<input type="hidden" id="isMustFields" name="isMustFields" value="<%= isMustFields %>">
	<td colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="저장" onclick="frm_check();">
	</td>
</tr>
</table>
<form name="tfrom" method="post">
	<input type="hidden" name="itemid" value="<%= itemid %>">
	<input type="hidden" name="itemoption" value="<%= itemoption %>">
	<input type="hidden" name="attrs" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->