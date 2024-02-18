<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/nvstorefarm/nvstorefarmcls.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

Dim itemid, oNvstorefarm
itemid = requestCheckVar(request("itemid"),10)
If (itemid = "") Then
    itemid = -1
End If
Set oNvstorefarm = new CNvstorefarm
	oNvstorefarm.FRectItemID = itemid
	oNvstorefarm.getShoppingWindowImageList
Dim i
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript'>

// ============================================================================
// 저장하기
function SubmitSave() {

    if (validate(itemreg)==false) {
        return;
    }

    if (itemreg.imgadd1.value == "del") {
        alert("기본이미지는 필수입니다.");
        return;
    } else {
        if (itemreg.imgadd1.value != "") {
            if (CheckImage(itemreg.imgadd1, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
                return;
            }
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage(itemreg.imgadd2, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage(itemreg.imgadd3, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd4.value != "") {
        if (CheckImage(itemreg.imgadd4, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage(itemreg.imgadd5, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
            return;
        }
    }

    if(confirm("저장하시겠습니까?") == true){
        itemreg.submit();
    }
}

function CloseWindow() {
    // opener.focus();
    window.close();
}


// ============================================================================
// 이미지표시
function ClearImage(img,fsize,wd,ht) {
    document.getElementById("div"+ img.name).style.display = "none";
	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "del";
}

function CheckExtension(imgname, allowext) {
    var ext = imgname.lastIndexOf(".");
    if (ext < 0) {
        return false;
    }

    ext = imgname.toLowerCase().substring(ext + 1);
    allowext = "," + allowext + ",";
    if (allowext.indexOf(ext) < 0) {
        return false;
    }

    return true;
}

function CheckImage(img, filesize, imagewidth, imageheight, extname, fsize)
{
    var ext;
    var filename;

	filename = img.value;
	if (img.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("이미지파일은 다음의 파일만 사용하세요.[" + extname + "]");
        ClearImage(img,fsize, imagewidth, imageheight);
        return false;
    }

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "";

    return true;
}

function CheckImage2(img, filesize, imagewidth, imageheight, extname, fsize, num)
{
    var ext;
    var filename;
    var imgcnt = $('input[name="addimgname"]').length;

	filename = img.value;
	if (img.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("이미지파일은 다음의 파일만 사용하세요.[" + extname + "]");
        ClearImage2(img,fsize, imagewidth, imageheight, num);
        return false;
    }

	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "";
    }else{
    	document.itemreg.addimgdel.value = "";
    }

    return true;
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/nvstorefarmImageModify.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="itemid" value="<%= itemid %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oNvstorefarm.GetImageByIdx(1) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%= oNvstorefarm.GetImageByIdx(1) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd1" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (450X450,jpg)
	  <input type="hidden" name="add1" />
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oNvstorefarm.GetImageByIdx(2) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=oNvstorefarm.GetImageByIdx(2) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (450X450,jpg)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oNvstorefarm.GetImageByIdx(3) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=oNvstorefarm.GetImageByIdx(3) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd3" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (450X450,jpg)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oNvstorefarm.GetImageByIdx(4) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=oNvstorefarm.GetImageByIdx(4) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd4" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (450X450,jpg)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oNvstorefarm.GetImageByIdx(5) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=oNvstorefarm.GetImageByIdx(5) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd5" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40);" class="text" size="40">
		<input type="button" value="이미지지우기" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (450X450,jpg)
		<input type="hidden" name="add5">
	</td>
</tr>
</table>
</form>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="저장하기" onClick="SubmitSave()">
          <input type="button" value="취소하기" onClick="CloseWindow()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<%
set oNvstorefarm = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->