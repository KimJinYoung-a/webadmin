<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Response.CharSet = "euc-kr"
%>
<script language='javascript'>
function sendDispCateItem() {
	if( $("#cate").val() == "" ){
		alert('카테고리를 선택하세요');
		return false;
	}
	document.frm.target = "xLink";
	document.frm.submit();
}
function jsCateCodeSelectBox(cdl){
	$.ajax({
		url: "/admin/etc/coupang/cate_selectbox_ajax.asp?cdl="+cdl,
		cache: false,
		success: function(message) {
			$("#categoryselectbox_a").empty().append(message);
		}
	});
}
</script>
<%
Dim oCoupang, i
Set oCoupang = new CCoupang
	oCoupang.FCurrPage	= 1
	oCoupang.FPageSize	= 50
	oCoupang.getCateLargeList

	If oCoupang.FResultCount > 0 Then
		Response.Write "<select id=""cate"" name=""cdl"" class=""select"" onchange=""jsCateCodeSelectBox(this.value);"">" & vbCrLf
		Response.Write "<option value="""">1 Depth</option>" & vbCrLf
		For i=0 To oCoupang.FResultCount-1
			Response.Write "<option value=""" & oCoupang.FItemList(i).FCode_large & """>" & oCoupang.FItemList(i).FCode_nm &"</option>"
		Next
		Response.Write "</select>"
	End If
Response.Write "<span id='categoryselectbox_a'>"
Response.Write "</span>"
set oCoupang = Nothing
%>
<input type="button" value="추가" class="button" onclick="sendDispCateItem()" />
<input type="button" value="취소" class="button" onclick="$('#lyrCateAdd').fadeOut();" />
<!-- #include virtual="/lib/db/dbclose.asp" -->