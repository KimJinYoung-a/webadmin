<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Response.CharSet = "euc-kr"
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function sendDispCateItem() {
	if( $("#cate").val() == "" ){
		alert('ī�װ��� �����ϼ���');
		return false;
	}
	if( $("#cate2").val() == "" ){
		alert('ī�װ��� �����ϼ���');
		return false;
	}
	document.frm.target = "xLink";
	document.frm.submit();
}
function jsCateCodeSelectBox(c, d){
	$.ajax({
		url: "/admin/etc/ssg/act_CategorySelectMulti.asp?depth="+d+"&cate="+c,
		cache: false,
		success: function(message) {
			$("#categoryselectbox_a").empty().append(message);
		}
	});
}
</script>
<%
Dim oSsg, i
Set oSsg = new Cssg
	oSsg.FCurrPage	= 1
	oSsg.FPageSize	= 50
	oSsg.getCateLargeList

Response.Write "<span id='categoryselectbox_a'>"
	If oSsg.FResultCount > 0 Then
		Response.Write "<select id=""cate"" name=""cate"" class=""select"" onChange=""jsCateCodeSelectBox(this.value,2);"" >" & vbCrLf
		Response.Write "<option value="""">1 Depth</option>" & vbCrLf
		For i=0 To oSsg.FResultCount-1
			Response.Write "<option value=""" & oSsg.FItemList(i).FCode_large & """>" & oSsg.FItemList(i).FCode_nm &"</option>"
		Next
		Response.Write "</select>"
	End If
response.write "</span>"
set oSsg = Nothing
%>
<input type="button" value="�߰�" class="button" onclick="sendDispCateItem()" />
<input type="button" value="���" class="button" onclick="$('#lyrCateAdd').fadeOut();" />
<!-- #include virtual="/lib/db/dbclose.asp" -->