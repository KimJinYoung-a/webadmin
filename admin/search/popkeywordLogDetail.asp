<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<%
Dim idx, arrList, oKeyList, i, rowNum
idx		= request("idx")
SET oKeyList = new cItemContent
	oKeyList.FRectIdx = idx
	oKeyList.getKeyWordLogDetailList
	arrList = oKeyList.fnkeywordMaster(idx)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
</script>
<table width="100%">
<tr>
	<td align="right"><input type="button" value="목록" class="button" onclick="location.href='/admin/search/popkeywordLog.asp';"></td>
</tr>
</table>
<p />
<table width="100%">
<tr>
	<td align="LEFT"><strong>변경 이력 정보</strong></td>
	<td align="RIGHT">*목록에서 바로 변경 적용한 키워드 정보는 공란입니다.</td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">변경 구분</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT">
			<%
				Select Case arrList(1, 0)
					Case "I"		response.write "등록"
					Case "U"		response.write "수정"
					Case "D"		response.write "삭제"
				End Select
			%>
			</td>
			<td width="15%">변경 키워드</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT">
			<%
				If arrList(1, 0) = "U" Then
					If arrList(4, 0) <> "" Then
						response.write arrList(3, 0) & " → " & arrList(4, 0)
					Else
						response.write arrList(4, 0)
					End If
				Else
					response.write arrList(4, 0)
				End If
			%>	
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">변경자</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT"><%= arrList(8, 0) %></td>
			<td width="15%">변경일</td>
			<td width="35%" bgcolor="#FFFFFF" align="LEFT"><%= arrList(7, 0) %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">비고</td>
			<td colspan="3" bgcolor="#FFFFFF" align="LEFT"><%= arrList(5, 0) %></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="15%">변경 내용</td>
			<td colspan="3" bgcolor="#FFFFFF" align="LEFT"><%= arrList(2, 0) %></td>
		</tr>
	</td>
</tr>	
</table>
<br/>
<table width="100%">
<tr>
	<td align="LEFT"><strong>변경 상품 정보</strong></td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="30">번호</td>
			<td width="100">전시카테고리</td>
			<td width="50">상품코드</td>
			<td width="50">이미지</td>
			<td width="100">브랜드ID</td>
			<td width="250">상품명</td>
			<td>키워드</td>
		</tr>
	<%
		rowNum = oKeyList.FTotalcount
		For i = 0 to oKeyList.FResultCount - 1
	%>
		<tr align="center" bgcolor="#FFFFFF" height="30">
			<td><%= rowNum %></td>
			<td><%= oKeyList.FItemList(i).FCatename %></td>
			<td><%= oKeyList.FItemList(i).FItemid %></td>
			<td><img src="<%= oKeyList.FItemList(i).Fsmallimage %>" width="50"></td>
			<td><%= oKeyList.FItemList(i).FMakerid %></td>
			<td><%= oKeyList.FItemList(i).FItemname %></td>
			<td><%= oKeyList.FItemList(i).FKeywords %></td>
		</tr>
	<%
			rowNum = rowNum - 1 
		Next
	%>
		<tr align="center" bgcolor="#FFFFFF" height="30">
			<td colspan="7"><input type="button" class="button" value="닫기" onclick="self.close();"></td>
		</tr>			
		</table>
	</td>
</tr>
</table>
<% SET oKeyList = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->