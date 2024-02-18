<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/categorymaster/displaycate2/classes/displaycateCls.asp"-->

<%
	Response.CharSet = "euc-kr"

	Dim cDisp, cDispOld, i

	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 1
	cDisp.GetDispCateList()

	SET cDispOld = New cDispCate
	cDispOld.FCurrPage = 1
	cDispOld.FPageSize = 2000
	cDispOld.FRectDepth = 1
	cDispOld.GetOldDispCateList()
%>
<script>
$(function() {
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});

function jsSaveItemMove(){
	if($('input[name="catecode_a"]').val() == ""){
		alert("기본 카테고리를 선택하세요.");
		$('input[name="catecode_a"]').focus();
		return;
	}
	
	if($('input[name="catecode_b"]').val() == ""){
		alert("신규 카테고리를 선택하세요.");
		$('input[name="catecode_b"]').focus();
		return;
	}

	if(confirm("선택하신대로 카테고리를 지정하시겠습니까?") == true) {
		frmItemAllMove.submit();
	}
}

function jsCateCodeSelectBox(c,d,g){
	$.ajax({
			url: "display_cate_selectbox_ajax.asp?depth="+d+"&cate="+c+"&gubun="+g+"",
			cache: false,
			success: function(message)
			{
				if(g == "a"){
					$("#categoryselectbox_a").empty().append(message);
				}else{
					$("#categoryselectbox_b").empty().append(message);
				}
			}
	});
}

function jsOldCateCodeSelectBox(c,d,g){
	$.ajax({
			url: "display_oldcate_selectbox_ajax.asp?depth="+d+"&cate="+c+"&gubun="+g+"",
			cache: false,
			success: function(message)
			{
				if(g == "a"){
					$("#categoryselectbox_a").empty().append(message);
				}else{
					$("#categoryselectbox_b").empty().append(message);
				}
			}
	});
}
</script>

<input type="hidden" name="mode" value="getOldCate">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr>
	<td bgcolor="#F3F3FF" width="100" height="35"></td>
	<td bgcolor="#FFFFFF" align="center"><b>기존 카테고리에서 상품 지정</b></td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" align="center" height="35">기존 카테고리</td>
	<td bgcolor="#FFFFFF">
		<div id="categoryselectbox_a">
		<%
		If cDispOld.FResultCount > 0 Then
			Response.Write "<select name=""cate"" class=""select"" onChange=""jsOldCateCodeSelectBox(this.value,2,'a');"">" & vbCrLf
			Response.Write "<option value="""">1 Depth</option>" & vbCrLf
			For i=0 To cDispOld.FResultCount-1
				Response.Write "<option value=""" & cDispOld.FItemList(i).FCateCode & """>" & cDispOld.FItemList(i).FCateName & "</option>"
			Next
			Response.Write "</select>"
		End If
		%>
		<input type="hidden" name="catecode_a" value="">
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" align="center" height="35">신규 카테고리</td>
	<td bgcolor="#FFFFFF">
		<div id="categoryselectbox_b">
		<%
		If cDisp.FResultCount > 0 Then
			Response.Write "<select name=""cate"" class=""select"" onChange=""jsCateCodeSelectBox(this.value,2,'b');"">" & vbCrLf
			Response.Write "<option value="""">1 Depth</option>" & vbCrLf
			For i=0 To cDisp.FResultCount-1
				Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """>" & cDisp.FItemList(i).FCateName & "</option>"
			Next
			Response.Write "</select>"
		End If
		%>
		<input type="hidden" name="catecode_b" value="">
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="35"></td>
	<td bgcolor="#FFFFFF">
		<span class="rdoUsing">
			<input type="radio" name="onlythiscate" id="useyn_1" value="Y" checked /><label for="useyn_1">이동할 카테고리 상품만 이동</label>
			<input type="radio" name="onlythiscate" id="useyn_2" value="N" /><label for="useyn_2">이동할 카테고리 하위 뎁스 상품 전부 이동</label>
		</span>
	</td>
</tr>
<tr>
	<td id="lyrSbmBtn" bgcolor="#FFFFFF" colspan="2">
		<table width="100%" class="a">
		<tr>
			<td></td>
			<td align="right"><input type="button" value="이  동" onClick="jsSaveItemMove()"></td>
		</tr>
		</table>
		<script>
			$("#lyrSbmBtn input").button();
		</script>
	</td>
</tr>
</table>
<%
	SET cDisp = Nothing
	SET cDispOld = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->