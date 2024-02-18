<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/RedRibbon/redRibbonManagerCls.asp"-->
<%
	Dim vAction, vIdx, vPage, vAnniv, vUseYN, vContent, arrList, arrView, iTotCnt, cEvtList, intLoop
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	vAction 	= NullFillWith(Request("action"), "")
	vIdx 		= NullFillWith(Request("idx"), "")
	vAnniv		= NullFillWith(Request("anniv"), "")
	vUseYN		= NullFillWith(Request("useyn"), "")
	vContent	= NullFillWith(html2db(Request("contents")), "")
	vPage		= NullFillWith(Request("iC"), 1)
	iPerCnt 	= 10

	set cEvtList = new giftManagerCls
	cEvtList.FPageSize 	= 5
	cEvtList.FCurrPage	= 1
	cEvtList.FMstNo 	= vIdx
	cEvtList.FUseYN 	= vUseYN
	
	If vAction = "insert_proc" OR vAction = "update_proc" OR vAction = "delete_proc" Then
		cEvtList.FAnniv		= vAnniv
		cEvtList.FContent	= vContent
		If vAction = "insert_proc" Then
			cEvtList.FGubun	= "I"
		ElseIf vAction = "update_proc" Then
			cEvtList.FGubun	= "U"
		ElseIf vAction = "delete_proc" Then
			cEvtList.FGubun	= "D"
		End If
		arrList = cEvtList.getGiftServiceMainList
		Response.Write "<script>alert('ok!');location.href='Pop_Gift_Mng.asp';</script>"
	ElseIf vAction = "" Then
		vAction = "insert"
	End If
	
	If vIdx <> "" Then
		cEvtList.FGubun	= "V"
		arrView = cEvtList.getGiftServiceMainList
	 	IF isArray(arrView) THEN
	 		vAnniv 		= arrView(1,0)
	 		vUseYN 		= arrView(3,0)
	 		vContent 	= db2html(arrView(2,0))
	 	End If
	End If

	cEvtList.FGubun	= "L"
	cEvtList.FMstNo = ""
	cEvtList.FCurrPage	= ((vPage - 1) * 5 )+1
	arrList = cEvtList.getGiftServiceMainList
	iTotCnt = cEvtList.FTotCnt
	iTotalPage 	=  int((iTotCnt-1)/5) +1
	
 	set cEvtList = nothing
 	
%>
<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
function checkform(frm1)
{
	if (frm1.anniv.value == "")
	{
		frm1.anniv.focus();
		alert("기념일을 입력하세요!");
		return false;
	}
	if (frm1.contents.value == "")
	{
		alert("항목을 선택하세요!");
		frm1.contents.focus();
		return false;
	}
}
function go1(a,b)
{
	if (a == 'e')
	{
		location.href = "?idx="+b+"&action=update";
	}
	else if (a == 'd')
	{
		if(confirm("진짜 삭제하시겠습니까?") == true)
		{
			frm1.action.value = "delete_proc";
			frm1.idx.value = b;
			frm1.submit();
		}
		else
		{
			return false;
		}
	}
}
function jsGoPage(iP){
	location.href = "?iC="+iP+"";
}
</script>
<body bgcolor="#FFFFFF" text="#000000" topmargin="10" leftmargin="10">
<table width="480" border="0" cellpadding="0" cellspacing="0" class="a" style="padding-bottom:10;">
<tr>
	<td align="left"><b>Gift Service 관리</b></td>
	<td align="right"><input type="button" class="button" value="창닫기" onclick="window.close();"></td>
</tr>
</table>
<table width="480" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm1" method="post" action="Pop_Gift_Mng.asp" onSubmit="return checkform(this);">
<input type="hidden" name="action" value="<%=vAction%>_proc">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="iC" value="<%=vPage%>">
<tr height="50">
	<td width="60" bgcolor="#E6E6E6" align="center">기념일</td>
	<td bgcolor="#FFFFFF"><input type="text" size="12" name="anniv" value="<%=vAnniv%>"> 예)2009년 6월
	<br>* 밑줄부분 입력: 기념일이 <font color="red"><u>2009년 6월</u></font>인 고객님들만 사연을 남겨주세요
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">사용여부</td>
	<td bgcolor="#FFFFFF">
		<label for="useyn1"><input type="radio" name="useyn" id="useyn1" value="1" <% If vUseYN = True Then Response.Write "checked" End If %>> 사용</label>&nbsp;&nbsp;&nbsp;
		<label for="useyn2"><input type="radio" name="useyn" id="useyn2" value="0" <% If vUseYN = False Then Response.Write "checked" End If %>> 사용안함</label>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">내용</td>
	<td bgcolor="#FFFFFF">
		<textarea name="contents" class="input" rows="10" cols="65"><%=vContent%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" style="padding:5 5 5 5;" align="right"><input type="button" class="button" value="취 소" onClick="location.href='?iC=<%=vPage%>';">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" class="button" value="저 장"></td>
</tr>
</form>
</table>
<br>
<table border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="60" bgcolor="#E6E6E6" align="center">No.</td>
	<td width="120" bgcolor="#E6E6E6" align="center">기념일</td>
	<td width="80" bgcolor="#E6E6E6" align="center">사용여부</td>
	<td width="100" bgcolor="#E6E6E6" align="center"></td>
</tr>
	<%
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
	%>
		   	<tr align="center" height="30" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
		   		<td><%=arrList(0,intLoop)%></td>
		   		<td align="center"><%=arrList(1,intLoop)%></td>
		   		<td align="center"><% If arrList(3,intLoop) = True Then Response.Write "사용" Else Response.Write "사용안함" End If %></td>
		   		<td align="center">
		   			<input type="button" class="button" value="수정" onClick="go1('e','<%=arrList(0,intLoop)%>')">&nbsp;&nbsp;
		   			<input type="button" class="button" value="삭제" onClick="go1('d','<%=arrList(0,intLoop)%>')">
		   		</td>
		   	</tr>
	<%	
			Next
		ELSE
	%>
	   	<tr align="center">
	   		<td colspan="4" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
	   	</tr>
   <%END IF%>
</table>
<!-- 페이징처리 -->
<%
iStartPage = (Int((vPage-1)/iPerCnt)*iPerCnt) + 1

If (vPage mod iPerCnt) = 0 Then
	iEndPage = vPage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
<table width="390" border="0" cellpadding="0" cellspacing="0" class="a">
    <tr valign="bottom" height="25">        
        <td valign="bottom" align="center">
         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(vPage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
		<%
				end if
			next
		%>
    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
		<% else %>[next]<% end if %>
        </td>        
    </tr>    
    </form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->