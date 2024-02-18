<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  AppDailyColor등록 및 수정
' History : 2013.12.17 김진영 생성
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
<!-- #include virtual="/lib/classes/appmanage/appColorCls.asp" -->
<%
Dim ocolor, mode, i
Dim yyyymmdd
yyyymmdd	= request("yyyymmdd")

If yyyymmdd = "" Then
	mode = "I"
Else
	mode = "U"
End If

Set ocolor = new AppColorList
	ocolor.FRectyyyymmdd = yyyymmdd
	ocolor.GetSelectOneMasterColor
%>
<script language="javascript">
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/appmanage/pop_appColorList_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}
function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
function form_check(){
	var frm = document.frm;
	if(frm.yyyymmdd.value == ''){
		alert('오픈일을 입력하세요');
		frm.yyyymmdd.focus();
		return;
	}
	if(frm.ImageUrl.value == ''){
		alert('대표Image를 등록하세요');
		return;
	}
	if(frm.ImageUrl2.value == ''){
		alert('대표Image2를 등록하세요');
		return;
	}
	if(frm.color_idx.value == ''){
		alert('오픈할 색상을 선택하세요');
		frm.color_idx.focus();
		return;
	}
	frm.submit();
}
</script>
<table border="0" width="100%" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/appmanage/appDailyColor_process.asp" style="margin:0px;">
<input type="hidden" name="ImageUrl" value="<%= ocolor.FOneItem.FImageUrl %>">
<input type="hidden" name="ImageUrl2" value="<%= ocolor.FOneItem.FImageUrl2 %>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="<%= mode %>">
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">오픈할 날짜</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="yyyymmdd" class="text" size="10" onClick="jsPopCal('yyyymmdd');"  style="cursor:hand;" value="<%= ocolor.FOneItem.FYyyymmdd %>">
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">대표Image<br>2048x1536 (4:3)</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('colorDailyList','<%= ocolor.FOneItem.FImageUrl %>','ImageUrl','spanban3')" class="button">
		<div id="spanban3" style="padding: 5 5 5 5">
			<% IF ocolor.FOneItem.FImageUrl <> "" THEN %>
				<img src="<%= ocolor.FOneItem.FImageUrl %>" border="0">
				<a href="javascript:jsDelImg('ImageUrl','spanban3');"><img src="/images/icon_delete2.gif" border="0"></a>
			<% End If %>
		</div>
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">대표Image2<br>1920x1080 (16:9)</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('colorDailyList','<%= ocolor.FOneItem.FImageUrl2 %>','ImageUrl2','spanban4')" class="button">
		<div id="spanban4" style="padding: 5 5 5 5">
			<% IF ocolor.FOneItem.FImageUrl2 <> "" THEN %>
				<img src="<%= ocolor.FOneItem.FImageUrl2 %>" border="0">
				<a href="javascript:jsDelImg('ImageUrl2','spanban4');"><img src="/images/icon_delete2.gif" border="0"></a>
			<% End If %>
		</div>
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">당일 오픈할 색상</td>
	<td bgcolor="#FFFFFF">
		<%= RegedColorBox("color_idx", ocolor.FOneItem.Fcolor_idx) %>
	</td>
</tr>
<% If mode = "U" Then %>
<tr>
	<td align="center" bgcolor="#FFFFFF">상세상품</td>
	<td bgcolor="#FFFFFF">
		<iframe id="iframG" frameborder="0" width="100%" src="/admin/appmanage/iframe_appDailyColorDetail.asp?yyyymmdd=<%=yyyymmdd%>" height=500%></iframe>
	</td>
</tr>
<% else %>
<tr>
	<td align="center" bgcolor="#FFFFFF">상세상품</td>
	<td bgcolor="#FFFFFF">
		신규등록 완료후 상세상품을 입력 하실수 있습니다.
	</td>
</tr>
<% End If %>
<tr>
	<td align="center" bgcolor="#FFFFFF"colspan="2">
		<input type="button" value="목록으로" onclick="location.href='/admin/appmanage/appDailyColorList.asp?menupos=<%=menupos%>';" class="button">
		<input type="button" value="저장" onclick="form_check();" class="button">
	</td>
</tr>
</form>
</table>
<% Set ocolor = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->