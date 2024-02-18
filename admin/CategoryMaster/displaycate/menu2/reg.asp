<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMenuCls.asp"-->

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<Script>
function checkform(f){
	if(f.disp1.value == ""){
		alert("카테고리를 선택하세요.");
		return false;
	}
	if(f.type.value == ""){
		alert("구분을 선택하세요.");
		return false;
	}
	if(f.type.value == "issue_image" && f.itemid.value == ""){
		alert("구분이 issue_image 인 경우 상품코드를 입력해야합니다.");
		f.itemid.focus();
		return false;
	}
	if(f.subject.value == ""){
		alert("텍스트를 입력하세요.");
		f.subject.focus();
		return false;
	}
	if(f.linkurl.value == ""){
		alert("링크를 입력하세요.");
		f.linkurl.focus();
		return false;
	}
	if(f.sortno.value == ""){
		alert("정렬번호를 입력하세요.");
		f.sortno.focus();
		return false;
	}
	return true;
}
</script>

<%
	Dim cMenu, i, vArr, vIdx, vDisp1, vType, vSubject, vItemID, vLinkURL, vSDate, vEDate, vUseYN, vSortNo, vRegdate
	vIdx = Request("idx")
	vDisp1 = Request("disp1")
	If vSortNo = "" Then vSortNo = "99" End If
	If vUseYN = "" Then vUseYN = "y" End If
	If vIdx <> "" Then
		rsget.Open "select * from db_item.dbo.tbl_display_cate_menu_top where idx = '" & vIdx & "'",dbget,1
		vDisp1		= rsget("disp1")
		vType 		= rsget("type")
		vSubject	= db2html(rsget("subject"))
		vItemID		= rsget("itemid")
		If vItemID = "0" Then
			vItemID = ""
		End If
		vLinkURL	= rsget("linkurl")
		vSDate 		= rsget("sdate")
		vEDate 		= rsget("edate")
		vUseYN 		= rsget("useyn")
		vSortNo		= rsget("sortno")
		vRegdate 	= rsget("regdate")
		rsget.Close
	End IF
	
	Set cMenu = New cDispCateMenu
	vArr = cMenu.GetDispCate1Depth()
	Set cMenu = Nothing
%>

<form name="frm1" action="proc.asp" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="idx" value="<%=vIdx%>">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% If vIdx <> "" Then %>
<tr bgcolor="#FFFFFF">
	<td>idx</td>
	<td><%=vIdx%> (등록일:<%=vRegdate%>)</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
	<td>카테고리</td>
	<td>
		<select name="disp1" class="select">
		<option value="">-카테고리선택-</option>
		<%
			For i=0 To UBound(vArr,2)
				Response.Write "<option value='" & vArr(0,i) & "' " & CHKIIF(CStr(vDisp1)=CStr(vArr(0,i)),"selected","") & ">" & vArr(1,i) & "</option>" & vbCrLf
			Next
		%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>구분</td>
	<td>
		<select name="type" class="select">
			<option value="">-구분선택-</option>
			<option value="issue_text" <%=CHKIIF(vType="issue_text","selected","")%>>issue_text</option>
			<option value="issue_image" <%=CHKIIF(vType="issue_image","selected","")%>>issue_image</option>
		</select>	
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>상품코드</td>
	<td>
		<input type="text" name="itemid" class="text" value="<%=vItemID%>"> * 구분이 issue_image 인 경우 반드시 입력!
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>텍스트</td>
	<td>
		<input type="text" name="subject" class="text" value="<%=vSubject%>" maxlength="48" size="60">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>링크</td>
	<td>
		<input type="text" name="linkurl" class="text" value="<%=vLinkURL%>" size="60">
		<br>* http://www.10x10.co.kr 은 빼고 / 부터 입력하세요.
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td></td>
	<td>
		<b>시작일, 종료일은 단순히 날짜관리하기 위한 날짜입니다. 절대 반영일이 아닙니다.</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>시작일</td>
	<td>
		<input id="sdate" name="sdate" value="<%=vSDate%>" class="text" size="10" maxlength="10" readonly />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "sdate", trigger    : "sdate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>종료일</td>
	<td>
		<input id="edate" name="edate" value="<%=vEDate%>" class="text" size="10" maxlength="10" readonly />
		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_End = new Calendar({
					inputField : "edate", trigger    : "edate_trigger",
					onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>정렬번호</td>
	<td>
		<input type="text" name="sortno" class="text" value="<%=vSortNo%>" size="5">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>사용여부</td>
	<td>
		<select name="useyn" class="select">
			<option value="y" <%=CHKIIF(vUseYN="y","selected","")%>>사용</option>
			<option value="n" <%=CHKIIF(vUseYN="n","selected","")%>>사용안함</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td></td>
	<td style="padding-top:20px;">
		<input type="submit" value="저  장" style="height:30px;">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->