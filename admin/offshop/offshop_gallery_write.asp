<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 매장 Gallery
' Hieditor : 2007.01.01 서동석 생성
'			 2016.12.28 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_galleryCls.asp" -->

<%
Dim idx, vShopID, vImageURL, vUseYN, offnews, vMainYN
	idx = getNumeric(requestcheckvar(request("idx"),10))

If idx <> "" Then
	set offnews = New COffshopGallery
		offnews.FIdx = idx
		offnews.GetOffshopGalleryView
		
		vShopID		= offnews.FItemOne.FShopID
		vImageURL	= offnews.FItemOne.FImageURL
		vUseYN		= offnews.FItemOne.FUseYN
		vMainYN		= offnews.FItemOne.FMainYN
	set offnews = Nothing
End If
%>

<script type="text/javascript">

function SubmitForm(){
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

<form method="post" name="f" action="<%= uploadImgUrl %>/linkweb/offshop/offshop_gallery_act.asp" onsubmit="return false" enctype="multipart/form-data">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="file2" value="<%=vImageURL%>">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<input type="hidden" name="incompany" value="o">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">매장</td>
	<td align="left">
		<% drawSelectBoxOffShopdiv_New "shopid", vShopID, "1,3", "", " onClick='reg("""");'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">사진</td>
	<td align="left">
		<input type="file" name="file1" size="50" class="input_b">

		<% If idx <> "" Then %>
			<br><img src="<%=vImageURL%>">
		<% End If %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">사용여부</td>
	<td align="left">
		<input type="radio" name="useyn" value="Y" <% If vUseYN = "Y" OR vUseYN = "" Then %>checked<% End If %>> Y&nbsp;&nbsp;&nbsp;
		<input type="radio" name="useyn" value="N" <% If vUseYN = "N" Then %>checked<% End If %>> N
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td width="100" bgcolor="<%= adminColor("gray") %>">메인 사용여부</td>
	<td align="left">
		<input type="radio" name="mainyn" value="Y" <% If vMainYN = "Y" OR vMainYN = "" Then %>checked<% End If %>> Y&nbsp;&nbsp;&nbsp;
		<input type="radio" name="mainyn" value="N" <% If vMainYN = "N" Then %>checked<% End If %>> N
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="2">
		<input type="button" value="저장" onclick="SubmitForm();" class="button">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->