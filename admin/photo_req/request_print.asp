<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<%
'####################################################
' Description : 촬영 요청 수정 & 뷰 페이지
' History : 2012.03.15 김진영 생성
'####################################################

	Dim gub, gubnm, i, udate
	Dim cPhotoreq, rno, arrFileList, sMode2
	Dim PhotoCnt
	
	rno = request("req_no")
	gub = request("gub")

	set cPhotoreq = new Photoreq
		cPhotoreq.FReq_no = rno
		cPhotoreq.fnPhotoreqUpdate
	If cPhotoreq.FPhotoreqList(0).FReq_use = "" Then
		Call Alert_move("해당 정보가 없습니다","request_list.asp")
	End If
%>
<script language="Javascript">
<!--
function printpage() {
	window.print();
}
//-->
</script>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="25">
    <td align="left"><font size="3"><strong>촬영요청서_<%=cPhotoreq.FPhotoreqList(0).FReq_department%> / 
<%
		If isnull(cPhotoreq.FPhotoreqList(0).FMDid) = "False" Then
			response.write cPhotoreq.FPhotoreqList(0).FMDid&"("& cPhotoreq.FPhotoreqList(0).FReq_name &")"
		ElseIf isnull(cPhotoreq.FPhotoreqList(0).FMDid) = "True" or (cPhotoreq.FPhotoreqList(0).FMDid) = "00" Then
			response.write cPhotoreq.FPhotoreqList(0).FReq_name
		End If
%>
	   	/ <%=cPhotoreq.FPhotoreqList(0).FPrd_name%></strong>
    </td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="25">
    <td align="right">중요도 : <%=cPhotoreq.FPhotoreqList(0).FImport_level%></td>
</tr>
</table>
<p>
<table width="100%" align="center"  border="1" bordercolordark="White" bordercolorlight="black" class="a" cellpadding="0" cellspacing="0">
<col width="15%"></col>
<col width="30%"></col>
<col width="17%"></col>
<col width="38%"></col>
<tr>
	<td height="30" bgcolor="#DDDDFF">촬영요청일시</td>
	<td><%=Left(cPhotoreq.FPhotoreqList(0).FReq_regdate,10)%></td>
	<td bgcolor="#DDDDFF">촬영요청자</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_name%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">촬영확정일시</td>
	<td colspan="3">
<%
	If IsNull(cPhotoreq.FPhotoreqList(0).FStart_date) Then
		response.write "&nbsp;"
	Else	
		For i = 0 to cPhotoreq.FResultcount -1
%>
				<font color="BLUE">시작 : <%=cPhotoreq.FPhotoreqList(i).FStart_date%></font> ~ <font color="RED">종료 : <%=cPhotoreq.FPhotoreqList(i).FEnd_date%></font><br>
<%
		Next
	End If
%>
	</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">포토그래퍼</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_photoname%>&nbsp;</td>
	<td bgcolor="#DDDDFF">스타일리스트</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FStylistname%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">촬영구분</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_gubun%></td>
	<td bgcolor="#DDDDFF">촬영용도</td>
	<td>
		<%=cPhotoreq.FPhotoreqList(0).FReq_use%>
		<%
			If cPhotoreq.FPhotoreqList(0).FReq_use_detail <> "" Then
				response.write "("&cPhotoreq.FPhotoreqList(0).FReq_use_detail&")"
			End If
		%>
	</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">상품명(기획전명)</td>
	<td colspan="3"><%=cPhotoreq.FPhotoreqList(0).FPrd_name%></td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">상품군/종</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FPrd_type%></td>
	<td bgcolor="#DDDDFF">판매가</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FPrd_price&"원"%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">브랜드ID</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FMakerid%>&nbsp;</td>
	<td bgcolor="#DDDDFF">요청부서/카테고리</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_department%>/<%=cPhotoreq.FPhotoreqList(0).FReq_codenm%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">필요 촬영군</td>
	<td colspan="3"><% call CheckBoxUseType("doc_use_type", rno, "3") %>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">메인 촬영 컨셉</td>
	<td colspan="3"><% call CheckBoxUseType("doc_use_concept", rno, "4") %>&nbsp;</td>
</tr>
</table>
<Br>
<table width="100%" align="center"  border="1" bordercolordark="White" bordercolorlight="black" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td height="30" bgcolor="#DDDDFF">상품 특징 및 주요 전달 사항</td>
</tr>
<tr>
	<td height="30" valign="top"><%=replace(cPhotoreq.FPhotoreqList(0).FReq_etc1,vbCrLf,"<br>")%></td>
</tr>
</table><br>
<p>
관련 링크 및 참고 url : <a href="<%=cPhotoreq.FPhotoreqList(0).FReq_url%>" target="_blank"><%=cPhotoreq.FPhotoreqList(0).FReq_url%></a>
<p><br>
<table width="100%" align="center" border="1" bordercolordark="White" bordercolorlight="black" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td height="30" bgcolor="#DDDDFF">촬영시 유의사항</td>
</tr>
<tr>
	<td height="30" valign="top"><%=replace(cPhotoreq.FPhotoreqList(0).FReq_etc2,vbCrLf,"<br>")%>&nbsp;</td>
</tr>
</table>
<p>
<table width="100%">
<tr>
	<td align="right">
		<input type="button" class="button_s" value="인쇄하기" onClick="printpage();">
	</td>
</tr>
</table>
</body>
<%set cPhotoreq = nothing%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->