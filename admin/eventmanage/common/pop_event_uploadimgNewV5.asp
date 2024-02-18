<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/pop_event_uploadimg.asp
' Description :  이벤트 이미지 등록
' History : 2007.02.22 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, vYear
Dim sOpt, wid, hei
sFolder = Request.Querystring("sF") 
sImg = Request.Querystring("sImg")
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	
sName = Request.Querystring("sName")
sSpan = Request.Querystring("sSpan")
sOpt = Request.Querystring("sOpt")
wid = Request.Querystring("wid")
hei = Request.Querystring("hei")

vYear = Request("yr")

If sSpan="spangift_img1" Then
wid = 170
hei = 170
End If
%>
<script language="JavaScript" src="https://code.jquery.com/jquery-3.2.1.min.js"></script>
<script language="javascript">
<!--
	function jsUpload(){
		if(document.frmImg.checkyn.value=="N"){
			alert("이미지 사이즈를 확인하고 올려주세요.");
			return false;
		}
		if(!document.frmImg.sfImg.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");			
			return false;
		}
		document.all.dvLoad.style.display = "";
	}

$(document).ready(function(){
	var _URL = window.URL || window.webkitURL;
	$('#sfImg').change(function (){
		var file = $(this)[0].files[0];
		img = new Image();
		var imgwidth = 0;
		var imgheight = 0;
		var maxwidth = <%=wid%>;
		var maxheight = <%=hei%>;

		img.src = _URL.createObjectURL(file);
		img.onload = function() {
			imgwidth = this.width;
			imgheight = this.height;

			$("#width").text(imgwidth);
			$("#height").text(imgheight);
			if(maxwidth==780 && maxheight==500)
			{
				if(imgwidth!=780 && imgheight!=500){
					alert("가로 사이즈 " + maxwidth + "px 세로 사이즈 " + maxheight + "px 에 맞춰 올려주세요!");
					document.frmImg.checkyn.value="N";
				}
			}
			else{
				if(imgwidth > maxwidth){
					alert("가로 사이즈 " + maxwidth + "px 이하로 올려주세요!");
					document.frmImg.checkyn.value="N";
				}
				if(imgheight > maxheight){
					alert("세로 사이즈 " + maxheight + "px 이하로 올려주세요!");
					document.frmImg.checkyn.value="N";
				}
			}
		}
	});
});


//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이미지 업로드 처리</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/event_admin/V5/event_upload_new.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="sOpt" value="<%=sOpt%>">
<input type="hidden" name="checkyn" value="Y">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지명</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg" id="sfImg"></td>
	</tr>	
	<%IF sImg <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName%></td>
	</tr>	
	<%END IF%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			+ 최대 파일사이즈 5MB(5,120KB) 이하만,<br>
			+ gif,jpg,png 타입의 파일만 등록가능
		</td>
	</tr>
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<div id="dvLoad" style="display:none;top:50px;left:20;position:absolute;background-color:gray;">
	<table border="0" class="a" cellpadding="5" cellspacing="5">
		<tr>
			<td> <font color="#FFFFFF">업로드 처리중입니다. 잠시만 기다려주세요~~</font></td>
		</tr>
	</table>
</div>