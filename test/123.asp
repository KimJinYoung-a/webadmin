<%@ language=vbscript %>
<% option explicit %>

<%
dim subtotalprice		: subtotalprice			= request.form("good_mny")

dim tt
'tt = CLNG(subtotalprice)

dim dispCate
response.write "tt="&tt
%>

<html>
<head>
	<title>제목1</title>
	<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
</head>
<body>

<table border=1 width='100%'>

<form method="post" name="frm" action="/test/123.asp">
<tr>
	<td>
		<input type="text" name="good_mny" value="">

	</td>
</tr>
<tr>
	<td>
		전시카테고리:

<script type="text/javascript">
$(function(){
	chgDispCate('<%=dispCate%>');
});

function chgDispCate(dc) {
    setTimeout( function() {  //
    	$.ajax({
    		url: "/common/module/dispCateSelectBox_response.asp?disp="+dc,
    		cache: false,
    		async: false,
    		success: function(message) {
           		// 내용 넣기
           		$("#lyrDispCtBox").empty().html(message);
           		$("#oDispCate").val(dc);
    		}
    	});
    }, 50);
}
</script>
<span id="lyrDispCtBox"></span>
<input type="hidden" name="disp" id="oDispCate" value="<%=dispCate%>">


	</td>
</tr>


<tr>
	<td colspan=2>
		<input type="button" value="검색" onclick="frm.submit();">
	</td>
</tr>
</form>
</table>

</body>
</html>

