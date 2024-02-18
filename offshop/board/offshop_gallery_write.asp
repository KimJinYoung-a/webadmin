<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_galleryCls.asp" -->

<%
	Dim idx, vShopID, vImageURL, vUseYN
	idx = request("idx")
	vShopID = fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))
	If idx <> "" Then
		dim offnews
		set offnews = New COffshopGallery
		offnews.FIdx = idx
		offnews.GetOffshopGalleryView
		
		vShopID		= offnews.FShopID
		vImageURL	= offnews.FImageURL
		vUseYN		= offnews.FUseYN
		set offnews = Nothing
	End If

%>

<script>
function SubmitForm()
{
    if (document.f.shopid.value == "") {
            alert("샵명을 선택하세요.");
            return;
    }
    
    <% If idx = "" Then %>
    if (document.f.file1.value == "") {
            alert("파일을 선택하세요.");
            return;
    }
    <% End IF %>

    if (confirm('저장 하시겠습니까?')){
        document.f.submit();
    }
}
</script>

<table  border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0" width="650" class="a">
<form method="post" name="f" action="<%= uploadImgUrl %>/linkweb/offshop/offshop_gallery_act.asp" onsubmit="return false" enctype="multipart/form-data">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="file2" value="<%=vImageURL%>">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<input type="hidden" name="shopid" value="<%=vShopID%>">
<input type="hidden" name="incompany" value="x">
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">첨부사진</td>
		<td bgcolor="white" style="padding:0">
				<input type="file" name="file1" size="50" class="input_b"><% If idx <> "" Then %><br>현재 이미지 <img src="<%=vImageURL%>" width="100"><% End If %>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">사용여부</td>
		<td bgcolor="white" style="padding:0">
			<input type="radio" name="useyn" value="Y" <% If vUseYN = "Y" OR vUseYN = "" Then %>checked<% End If %>> Y&nbsp;&nbsp;&nbsp;
			<input type="radio" name="useyn" value="N" <% If vUseYN = "N" Then %>checked<% End If %>> N
		</td>
	</tr>
	<tr>
		<td style="padding:0" colspan="2" align="right" bgcolor="white">
			<input type="button" value="Save" onclick="SubmitForm()" style="background-color:#dddddd; height:25; border:1 solid buttonface">&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</form>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->