<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 카테고리
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Response.CharSet = "euc-kr"

Dim cntDefault
cntDefault = requestCheckVar(Request("isDft"),10)
%>
<script>
function jsCateCodeSelectBox(c,d,g){
	$.ajax({
		url: "/admin/CategoryMaster/displaycate/display_cate_selectbox_ajax.asp?depth="+d+"&cate="+c+"&gubun="+g+"",
		cache: false,
		success: function(message) {
			$("#categoryselectbox_a").empty().append(message);
		}
	});
}

function sendDispCateItem() {
	var dcd,cnm,div,dpt;

	dcd = $("input[name='catecode_a']").val();
	cnm = $("select[name='cate'] option:selected");
	//div = $("select[name='selDispCateDiv'] option:selected").val();
	div = $("input[name='selDispCateDiv']").val();
	dpt = dcd.length/3
	if(dpt==0) {
		alert('카테고리를 선택해주세요.');
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
cDisp.GetDispCateList()

Response.Write "<span id='categoryselectbox_a'>"

If cDisp.FResultCount > 0 Then
	Response.Write "<select name=""cate"" class=""select"" onChange=""jsCateCodeSelectBox(this.value,2,'a');"">" & vbCrLf
	Response.Write "<option value="""">1 Depth</option>" & vbCrLf
	For i=0 To cDisp.FResultCount-1
		Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """>" & cDisp.FItemList(i).FCateName & chkIIF(cDisp.FItemList(i).FUseYN="N"," (사용안함)","") & "</option>"
	Next
	Response.Write "</select>"
End If

Response.Write "<input type='hidden' name='catecode_a' value='' />" &_
				"</span>"

set cDisp = Nothing
%>
<!--/ <select name="selDispCateDiv" class="select">
<option value="y" <%=chkIIF(cntDefault>0,"","selected")%>>기본</option>
<option value="n" <%=chkIIF(cntDefault>0,"selected","")%>>추가</option>
</select>-->
<%
	if cntDefault>0 then
		Response.Write "<input type=""hidden"" name=""selDispCateDiv"" value=""n"">"
	else
		Response.Write "<input type=""hidden"" name=""selDispCateDiv"" value=""y"">"
	end if
%>
<input type="button" value="추가" class="button" onclick="sendDispCateItem()" />
<input type="button" value="취소" class="button" onclick="$('#lyrDispCateAdd').fadeOut();" />
<!-- #include virtual="/lib/db/dbclose.asp" -->