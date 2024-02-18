<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<%
Response.CharSet = "euc-kr"

Dim cntDefault, chkCnt
cntDefault = RequestCheckvar(Request("isDft"),10)
chkCnt = RequestCheckvar(Request("chk"),10)
%>
<script>
function jsCateCodeSelectBox(c,d,g){
	$.ajax({
		url: "/academy/CategoryMaster/displaycate/display_cate_selectbox_ajax_upche.asp?depth="+d+"&cate="+c+"&gubun="+g+"",
		cache: false,
		success: function(message) {
			$("#categoryselectbox_a").empty().append(message);
		}
	});
}

function sendDispCateItem() {
	var dcd,cnm,div,dpt,catecode_depth;

	dcd = $("input[name='catecode_a']").val();
	cnm = $("select[name='cate'] option:selected");
	//div = $("select[name='selDispCateDiv'] option:selected").val();
	div = $("input[name='selDispCateDiv']").val();
	catecode_depth = $("input[name='catecode_depth']").val();
	dpt = dcd.length/3
	if(dpt!=catecode_depth || dpt == 0) {
		alert('ī�װ����� ������ depth���� �������ּ���.');
		return;
	}

	addDispCateItem(dcd,cnm,div,dpt);
}
</script>
<%
Dim cDisp, i
SET cDisp = New cDispCate
cDisp.FCurrPage = 1
cDisp.FPageSize = 2000
cDisp.FRectDepth = 1
cDisp.FRectUseYN = "Y"
cDisp.GetDispCateList()

Response.Write "<span id='categoryselectbox_a'>"

If cDisp.FResultCount > 0 Then
	Response.Write "<select name=""cate"" class=""select"" onChange=""jsCateCodeSelectBox(this.value,2,'a');"">" & vbCrLf
	Response.Write "<option value="""">1 Depth</option>" & vbCrLf
	For i=0 To cDisp.FResultCount-1
		Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """>" & cDisp.FItemList(i).FCateName & "</option>"
	Next
	Response.Write "</select>"
End If

Response.Write "<input type='hidden' name='catecode_a' value='' />" &_
				"<input type='hidden' name='catecode_depth' value='' />" &_
				"</span>"

set cDisp = Nothing

	if cntDefault>0 then
		Response.Write "<input type=""hidden"" name=""selDispCateDiv"" value=""n"">"
	else
		Response.Write "<input type=""hidden"" name=""selDispCateDiv"" value=""y"">"
	end if
%>
<% if chkCnt<2 then %><input type="button" value="�߰�" class="button" onclick="sendDispCateItem()" /><% end if %>
<input type="button" value="���" class="button" onclick="$('#lyrDispCateAdd').fadeOut();" />
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->