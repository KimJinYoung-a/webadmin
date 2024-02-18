<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
'###############################################
' Discription : 전시카테고리-상품속성 목록 Ajax
' History : 2013.08.06 허진원 : 신규 생성
'###############################################
Response.CharSet = "euc-kr"

'// 변수 선언
Dim dispCate
Dim oAttrib

'// 파라메터 접수
dispCate = request("dispcate")

'// 페이지정보 목록
	set oAttrib = new CAttrib
	oAttrib.FRectDispCate = dispCate
    oAttrib.GetAttribList4DispCate

	if oAttrib.FResultCount>0 then
%>
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="<%=chkIIF(dispCate="","add","modi")%>">
<table width="100%" cellpadding="2" cellspacing="2">
<tr>
	<td colspan="2" bgcolor="#F8F8F8">
		카테고리 :
		<span id="categoryselectbox_b">
		<%
		'//전시카테고리
		Dim cDisp, i
		
		if dispCate="" then
			SET cDisp = New cDispCate
			cDisp.FCurrPage = 1
			cDisp.FPageSize = 2000
			cDisp.FRectDepth = 1
			cDisp.GetDispCateList()
	
			If cDisp.FResultCount > 0 Then
				Response.Write "<select name=""catedsp"" class=""select"" onChange=""jsCateCodeSelectBox(this.value,2,'b');"">" & vbCrLf
				Response.Write "<option value="""">1 Depth</option>" & vbCrLf
				For i=0 To cDisp.FResultCount-1
					Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """>" & cDisp.FItemList(i).FCateName & "</option>"
				Next
				Response.Write "</select>"
			End If

			set cDisp = Nothing
		else
			response.Write getDispCateHistory(dispCate)
		end if

		
		%>
		<input type="hidden" name="catecode_b" value="<%=dispCate%>">
		</span>
	</td>
</tr>
<%
	for i=0 to oAttrib.FResultCount-1
%>
<tr>
	<td width="30" align="center" bgcolor="#FFFFFF"><input type="checkbox" name="attribDiv" id="atrDiv<%=i%>" value="<%=oAttrib.FItemList(i).FattribDiv%>" <%=chkIIF(oAttrib.FItemList(i).FchkCate,"checked","")%> /></td>
	<td bgcolor="#FFFFFF"><label for="atrDiv<%=i%>"><%=oAttrib.FItemList(i).FattribDivName%></label></td>
</tr>
<%
	next
%>
<tr>
	<td colspan="2" align="center">
		<% if dispCate<>"" then %>
		<input type="button" value=" 삭 제 " class="button" onclick="deleteItem()"> &nbsp; &nbsp;
		<% end if %>
		<input type="button" value=" 취소 " class="button" onclick="resizeArea('left');$('#lyrRightList').empty().html('카테고리-상품속성 편집영역');"> &nbsp; &nbsp;
		<input type="button" value=" 저 장 " class="button" onclick="saveItem()">
	</td>
</tr>
</table>
</form>
<% else %>
등록된 상품속성이 없습니다.<br>
"<a href='/admin/itemmaster/itemAttribute/itemAttribute_List.asp?menupos=1587'>[ON]상품관리>>상품속성 관리</a>" 메뉴에서 상품속성을 등록해주세요.
<%
	end if
	set oAttrib = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->