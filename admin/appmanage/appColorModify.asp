<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  AppColor 등록 및 수정
' History : 2013.12.16 김진영 생성
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/appColorCls.asp" -->
<%
Dim ocolor, mode, menupos, newColorCd
Dim i
Dim wcolorCode

wcolorCode	= request("wcolorCode")
menupos		= request("menupos")
newColorCd	= request("newColorCd")

Set ocolor = new AppColorList
	ocolor.FRectcolorCode = wcolorCode
	ocolor.GetSelectOneColor()
	If ocolor.FOneCount = 0 Then
		mode = "I"
	Else
		mode = "U"
	End If
%>
<script language="javascript">
function gotoColor(code){
	var frm = document.frmSearch;
	if(code == 'New'){
		frm.newColorCd.value='New';
		frm.submit();
	}else{
		frm.submit();
	}
}
function form_check(){
	var frm = document.frm;
	if(frm.colorName.value == ''){
		alert('색상명을 입력하세요');
		frm.colorName.focus();
		return;
	}
	frm.submit();
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
</script>
<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<form name="frmSearch" method="get" action="<%=CurrURL%>" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="newColorCd" value="">
<tr>
	<td>텐바이텐 웹에서 사용하는 색상들</td>
</tr>
<tr>
	<td>
		<table border="0" width="100%" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="30">
			<td bgcolor="#FFFFFF">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<% call ocolor.GetWebColorCode %>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<br>
<% If wcolorCode = "" AND newColorCd<>"New" Then %>
<input type="button" value="앱에서만 사용하는 색상등록" onclick="gotoColor('New');" class="button">
<% End If %>
<br><br>
<% If (newColorCd = "New") OR (wcolorCode <> "" AND newColorCd <> "New") Then %>
<%= ChkIIF(ocolor.FOneCount > 0, "<font color='RED'><strong>수정</strong></font>","<font color='BLUE'><strong>신규등록</strong></font>") %>
<table border="0" width="100%" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/appmanage/appColor_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%= ocolor.FOneItem.FIdx %>">
<input type="hidden" name="iconImage1" value="<%= ocolor.FOneItem.FIconImageUrl1 %>">
<input type="hidden" name="iconImage2" value="<%= ocolor.FOneItem.FIconImageUrl2 %>">
<%  If (newColorCd = "New") Then %>
<input type="hidden" name="colorCode" value="<%= newColorCd %>">
<% Else %>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">색상코드</td>
	<td bgcolor="#FFFFFF">
		<input type="text" size="3" readonly name="colorCode" value="<%= wcolorCode %>">
	</td>
</tr>
<% End If %>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">색상명</td>
	<td bgcolor="#FFFFFF">
		<input type="text"  name="colorName" value="<%= ocolor.FOneItem.FColorName %>">
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">IconImage1</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('colorList','<%= ocolor.FOneItem.FIconImageUrl1 %>','iconImage1','spanban')" class="button">
		<div id="spanban" style="padding: 5 5 5 5">
			<% IF ocolor.FOneItem.FIconImageUrl1 <> "" THEN %>
				<img src="<%=ocolor.FOneItem.FIconImageUrl1%>" border="0" width="50" height="50">
				<a href="javascript:jsDelImg('iconImage1','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
			<% END IF %>
		</div>
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">IconImage2</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnBan" value="이미지등록" onClick="jsSetImg('colorList','<%= ocolor.FOneItem.FIconImageUrl2 %>','iconImage2','spanban2')" class="button">
		<div id="spanban2" style="padding: 5 5 5 5">
			<% IF ocolor.FOneItem.FIconImageUrl2 <> "" THEN %>
				<img src="<%=ocolor.FOneItem.FIconImageUrl2%>" border="0" width="50" height="50">
				<a href="javascript:jsDelImg('iconImage2','spanban2');"><img src="/images/icon_delete2.gif" border="0"></a>
			<% END IF %>
		</div>
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">색상String</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="color_str" value="<%= ocolor.FOneItem.FColor_str %>">
		<font color="RED">ex) FFFFFF</font>
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">글자 RGB코드</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="word_rgbCode" value="<%= ocolor.FOneItem.FWord_rgbCode %>">
		<font color="RED">ex) FFFFFF</font>
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%= ChkIIF((ocolor.FOneItem.FIsusing = "") OR (ocolor.FOneItem.FIsusing = "Y"),"checked","") %> >Y
		<input type="radio" name="isusing" value="N" <%= ChkIIF(ocolor.FOneItem.FIsusing = "N","checked","") %> >N
	</td>
</tr>
<tr height="30">
	<td width="100" bgcolor="#FFFFFF">순서</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sortNo" value="<%= ChkIIF((ocolor.FOneItem.FSortNo=""),"0",ocolor.FOneItem.FSortNo) %>"></td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"colspan="2"><input type="button" value="저장" onclick="form_check();" class="button"></td>
</tr>
</form>
</table>
<% End If %>
<% Set ocolor = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->