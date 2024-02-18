<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  파트관리자 하위카테고리 담당자 리스트 폼
' History : 2011.01.25 김진영 생성
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/partpersonCls.asp"-->
<%
Dim idx, i, name
idx = requestCheckVar(request("idx"),10)
Dim plist, arlist3, arlist4
	Set plist = new Partlist
		plist.idx = idx
		arlist3 = plist.fnGetmolist2
		arlist4 = plist.fnGetmolist
	Set plist = nothing
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function modify(k,j){
	document.lform.action = "partcate_pop.asp?cc=<%=idx%>&idx="+k+"&mode=cmodify&sab="+j;
	document.lform.submit();
}
function hide(idx,sabun){
	if(confirm("확인버튼 클릭시 숨겨지게 됩니다. \n계속 하시겠습니까?")){
		document.lform.action = "partcate_proc.asp?idx="+idx+"&sabun="+sabun+"&mode=chide";
		document.lform.submit();
	}
}
function use(idx,sabun){
	if(confirm("확인버튼 클릭시 사용할 수 있게 됩니다. \n계속 하시겠습니까?")){
		document.lform.action = "partcate_proc.asp?idx="+idx+"&sabun="+sabun+"&mode=cuse";
		document.lform.submit();
	}
}
function updSordno(v){
	var sNo;
	sNo = $("#sortno_"+v).val();
	document.lform.action = "partcate_proc.asp?cc=<%=idx%>&idx="+v+"&sortNo="+sNo+"&mode=sortNo";
	document.lform.submit();
}

function updSordnoAll(){
	document.lform.action = "partcate_proc.asp";
	document.lform.mode.value ="sortNoAll";
	document.lform.submit();
}

$(function() {
	// 행 정렬
	$( "#mainList" ).sortable({
		placeholder: "ui-state-highlight",
		handle: '.handle',
		start: function(event, ui) {
			ui.placeholder.html('<td height="40" colspan="7" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
 			i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="margin:10px 0 5px 0;">
	<tr>
		<td style="text-align:left;">
			<input type="button" class="button" value="하위카테고리 등록" onclick="javascript:location.href='partcate_pop.asp?idx=<%=idx%>&mode=cinsert'">
		</td>
		<td style="text-align:right;">
			<input type="button" class="button_auth" value="순서 저장" onclick="updSordnoAll()">
		</td>
	</tr>
</table>
<form name="lform" method="post">
<input type="hidden" name="mode" value="" />
<input type="hidden" name="cc" value="<%=idx%>" />
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<% If IsEmpty(arlist4) = "False" Then %>
<input type = "hidden" name = "sname" value="<%= arlist4(1,0) %>">
<tr bgcolor="#DDDDFF">
	<td width="240" height="30">상위 카테고리 이름</td>
	<td colspan=6 bgcolor="#FFFFFF" height="30"><%= arlist4(1,0) %></td>
</tr>
<% end if %>
<tr>
	<td colspan="7"><b><center>하위 카테고리 담당자 데이터</center></b><p></td>
</tr>
<tbody id="mainList">
<%
Dim plist2, s
If IsEmpty(arlist3) = "False" Then
	Set plist2 = new Partlist
		plist2.idx = idx

	For i = 0 to Ubound(arlist3,2)
%>
<tr bgcolor="#DDDDFF" height="30">
	<td width="240" class="handle" style="cursor:grab;"><span class="ui-icon ui-icon-grip-solid-horizontal" style="display:inline-block;"></span> <%=arlist3(2,i) %></td>
	<td width="90" bgcolor="#FFFFFF"><%=arlist3(3,i)%><!--&nbsp;<%=arlist3(8,i)%>--></td>
	<td width="120" bgcolor="#FFFFFF"><%=arlist3(4,i)%></td>
	<td width="60" bgcolor="#FFFFFF" style="text-align:center;"><%=arlist3(5,i)%></td>
	<td width="40%" bgcolor="#FFFFFF"><%=arlist3(6,i)%></td>
	<td width="90" bgcolor="#FFFFFF" style="text-align:center;">
		<input type="hidden" name="idx" value="<%=arlist3(1,i) %>" />
		순서 : <input type="text" class="text" id="sortno_<%=arlist3(1,i) %>" size="1" name="sortno" value="<%=arlist3(10,i)%>">
	</td>
	<td width="120"  bgcolor="#FFFFFF" align="center">
		<a href="javascript:modify('<%=arlist3(1,i)%>','<%=arlist3(7,i)%>');"><img src="/images/icon_modify.gif" border="0"></a>
			<%If arlist3(11,i) = "Y" Then%><a href="javascript:hide('<%=idx%>','<%=arlist3(7,i)%>')"><img src="/images/icon_hide.gif" border="0"></a><%End If%>
			<%If arlist3(11,i) = "N" Then%><a href="javascript:use('<%=idx%>','<%=arlist3(7,i)%>')"><img src="/images/icon_use.gif" border="0"></a><%End If%>
	</td>
</tr>
<%
	Next
	Set plist2 = nothing
Else
%>
<tr bgcolor="#DDDDFF">
	<td bgcolor="#FFFFFF" height="30" colspan="7">목록이 없습니다</td>
</tr>
<%
End If
%>
</tbody>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->