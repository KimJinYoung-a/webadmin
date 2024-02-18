<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_main_pageItemcls.asp" -->

<%
Dim cDisp
dim cdl,idx,mode,i,did,page,divCd, cdm, vCateCode
vCateCode = Request("catecode") 
mode=request("mode")
idx=request("idx")
page=request("page")
divCd = request("divCd")
menupos=request("menupos")


'// 항목 구분 선택상자 서브함수 //
Sub DrawSelectBoxPageDiv(byval selectBoxName,selectedId)
   dim tmp_str1, tmp_str2, query1

	'# Select Box 생성
	tmp_str1 = "<select name='" & selectBoxName & "' onchange='chgDivSelect(this.value);'>" & vbCrLf

	'# OnChange 스크립트 생성
	tmp_str2 = "<script language='javascript'>" & vbCrLf &_
				"function chgDivSelect(dcd) { " & vbCrLf &_
				"	switch(dcd) { " & vbCrLf

	'#항목 구분 쿼리
	query1 = "Select divCd, divName, imgWidth, imgHeight, divType " &_
			"From [db_sitemaster].[dbo].tbl_category_mainItem_div " &_
			"Where isUsing = 'Y' " &_
			"Order by divCd Asc"
	rsget.Open query1,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
	rsget.Movefirst

		do until rsget.EOF
			tmp_str1 = tmp_str1 & "<option value='" & rsget("divCd") & "'"
			if Cstr(selectedId) = Cstr(rsget("divCd")) then tmp_str1 = tmp_str1 & " selected"
			tmp_str1 = tmp_str1 & ">[" & rsget("divCd") & "]" & rsget("divName") & "</option>" & vbCrLf
			
			tmp_str2 = tmp_str2 & "case '" & rsget("divCd") & "':" & vbCrLf &_
									"	runFormRowOnOff('" & rsget("divType") & "','" & rsget("imgWidth") & "','" & rsget("imgHeight") & "');" & vbCrLf &_
									"	break;" & vbCrLf

			rsget.MoveNext
		loop
	end if
	rsget.close
	tmp_str1 = tmp_str1 & "</select>" & vbCrLf

	tmp_str2 = tmp_str2 & "default:" & vbCrLf &_
						"	runFormRowOnOff('I','','');" & vbCrLf &_
						"	break;" & vbCrLf &_
						"}" & vbCrLf & "}" & vbCrLf & "</script>" & vbCrLf

	Response.Write tmp_str1
	Response.Write tmp_str2

end Sub
%>

<script language="javascript">
<!--
// 내용검사 / 실행
function subcheck(){
	var frm=document.inputfrm;

	if (frm.catecode.value.length<1) {
		alert('카테고리를 선택해 주세요..');
		frm.catecode.focus();
		return;
	}

	if (frm.divCd.value.length<1) {
		alert('항목구분을 선택해 주세요..');
		frm.divCd.focus();
		return;
	}
	
	if(confirm("입력한 내용으로 저장하시겠습니까?"))
		frm.submit()
	else
		return;
}

// 관련 상품 검색 팝업
function findItemId(frm)
{	 
	window.open("/common/pop_singleItemSelect.asp?disp="+document.inputfrm.catecode.value+"&itemid="+document.inputfrm.itemid.value+"&target=" + frm + "&ptype=","popSearch","width=1024,height=768,resizable=yes,scrollbars=yes,status=no,top=200,left=600");
}

// 폼 숨김/해제
function runFormRowOnOff(sw,iw,ih)
{
	var frm = document.all;
	switch(sw)
	{
		case "I":
			// 상품선택
			frm.row_item.style.display="";
			frm.row_img.style.display="none";
			frm.row_link.style.display="none";
			frm.imgSize.innerText = "";
			break;
		case "M":
			// 이미지선택
			frm.row_item.style.display="none";
			frm.row_img.style.display="";
			frm.row_link.style.display="";
			frm.imgSize.innerText = "(" + iw + "px × " + ih + "px)";
			break;
		case "B":
			// 상품 & 이미지선택
			frm.row_item.style.display="";
			frm.row_img.style.display="";
			frm.row_link.style.display="";
			frm.imgSize.innerText = "(" + iw + "px × " + ih + "px)";
			break;
	}
}

// 카테고리 변경시
function changecontent()
{
}
//-->
</script>
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="20">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top"><b>카테고리 메인 페이지 항목 등록/수정</b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="inputfrm" method="post" action="<%=uploadUrl%>/linkweb/doMainPageItem.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<% =mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<% if mode="add" then %>
<tr>
	<td width="100" bgcolor="#F0F0FD" align="center">카테고리선택</td>
	<td bgcolor="#FFFFFF">
	<%
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	'cDisp.FRectUseYN = "Y"
	cDisp.GetDispCateList()
	
	If cDisp.FResultCount > 0 Then
		Response.Write "<select name=""catecode"" class=""select"">" & vbCrLf
		Response.Write "<option value="""">선택</option>" & vbCrLf
		For i=0 To cDisp.FResultCount-1
			Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
		Next
		Response.Write "</select>&nbsp;&nbsp;&nbsp;"
	End If
	Set cDisp = Nothing
	%>
	</td>
</tr> 
<tr>
	<td width="100" bgcolor="#F0F0FD" align="center">항목선택</td>
	<td bgcolor="#FFFFFF"><% DrawSelectBoxPageDiv "divCd", divCd %></td>
</tr>
<tr name="row_item" id="row_item">
	<td align="center" bgcolor="#F0F0FD">상품번호</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemid" size="10">
		<input type="button" value="검색" onClick="findItemId('inputfrm')">
		<span name="itemname" id="itemname"></span>
	</td>
</tr>
<tr name="row_img" id="row_img" style="display:none">
	<td align="center" bgcolor="#F0F0FD">이미지 선택</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgFile" size="50">
		<span name="imgSize" id="imgSize"></span>
	</td>
</tr>
<tr name="row_link" id="row_link" style="display:none">
	<td align="center" bgcolor="#F0F0FD">링크 주소</td>
	<td bgcolor="#FFFFFF">
		http://www.10x10.co.kr <input type="text" name="linkURL" value="/" size="60">
		<br>※ 상대주소로 표기
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">정렬번호</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="sortno" size="10" value="0">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" checked>Y
		<input type="radio" name="isusing" value="N">N
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td colspan="2" align="center">
			<input type="button" value=" 저장 " onclick="subcheck();"> &nbsp;&nbsp;
			<input type="button" value=" 취소 " onclick="history.back();">
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New CateMainPage
	fmainitem.GetOnePageItem idx
%>
<tr>
	<td width="100" align="center" bgcolor="#F0F0FD">카테고리 선택</td>
	<td bgcolor="#FFFFFF">
	<%
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()
	
	If cDisp.FResultCount > 0 Then
		Response.Write "<select name=""catecode"" class=""select"">" & vbCrLf
		Response.Write "<option value="""">선택</option>" & vbCrLf
		For i=0 To cDisp.FResultCount-1
			Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(fmainitem.FDisp)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
		Next
		Response.Write "</select>&nbsp;&nbsp;&nbsp;"
	End If
	Set cDisp = Nothing
	%>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#F0F0FD" align="center">항목선택</td>
	<td bgcolor="#FFFFFF"><% DrawSelectBoxPageDiv "divCd", fmainitem.FdivCd %></td>
</tr>
<tr name="row_item" id="row_item" <% if fmainitem.FdivType="M" then %>style="display:none"<% end if %>>
	<td align="center" bgcolor="#F0F0FD">상품번호</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemid" size="10" value="<%=fmainitem.FitemId%>">
		<input type="button" value="검색" onClick="findItemId('inputfrm')">
		<span name="itemname" id="itemname"><%=fmainitem.Fitemname%></span>
	</td>
</tr>
<tr name="row_img" id="row_img" <% if fmainitem.FdivType="I" then %>style="display:none"<% end if %>>
	<td align="center" bgcolor="#F0F0FD">이미지 선택</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgFile" size="50">
		<span name="imgSize" id="imgSize">(<%=fmainitem.FimgWidth & "px × " & fmainitem.FimgHeight & "px"%>)</span>
		<br><%=fmainitem.FimgFile%>
	</td>
</tr>
<tr name="row_link" id="row_link" <% if fmainitem.FdivType="I" then %>style="display:none"<% end if %>>
	<td align="center" bgcolor="#F0F0FD">링크 주소</td>
	<td bgcolor="#FFFFFF">
		http://www.10x10.co.kr <input type="text" name="linkURL" value="<%=fmainitem.FlinkURL%>" size="60">
		<br>※ 상대주소로 표기
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">정렬번호</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="sortno" size="10" value="<%=fmainitem.FSortNo%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#F0F0FD">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%if fmainitem.FIsusing="Y" then response.write "checked" %> checked>Y
		<input type="radio" name="isusing" value="N" <%if fmainitem.FIsusing="N" then response.write "checked" %>>N
		<input type="hidden" name="orgUsing" value="<%=fmainitem.FIsusing%>">
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td colspan="2" align="center">
		<input type="button" value=" 저장 " onclick="subcheck();"> &nbsp;&nbsp;
		<input type="button" value=" 취소 " onclick="history.back();">
	</td>
</tr>
<% end if %>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
