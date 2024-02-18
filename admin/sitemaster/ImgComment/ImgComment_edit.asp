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
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/SitemasterClass/ImgCommentCls.asp"-->
<%
dim reviewid, mode

reviewid=request("reviewid")
mode =request("mode")

if mode="add" then reviewid=0

dim oitem
set oitem = new CItemImage

oitem.FRectReviewID = reviewid
oitem.GetOneItemImage

%>
<script language='javascript'>

function DelImage(src, imgname, hid){
	var imgcomp;
	imgcomp = eval("document." + imgname);
	if (hid.checked){
		imgcomp.src = '/images/space.gif';
	}
}

function jsSubmit(frm){

	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		frm.submit();
	}
}

function jsPopCal(sName){
		var winCal;
		winCal = window.open('/admin/sitemaster/imgComment/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
function showimage(img){
	var pop = window.open('/lib/showImage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
}
function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='ImgComment_image_input.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	//document.imginputfrm.maxFileSize.value = Fsize;
	//document.imginputfrm.maxFileWidth.value = Fwidth;
	//document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='<%= uploadImgUrl %>/linkweb/christmas/christmas_postcard_img_del.asp';
	document.imginputfrm.submit();
}

document.domain='10x10.co.kr';
</script>

<table width="700" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#B2B2B2">
<form name="frm_add" method="post" action="ImgComment_Proc.asp">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="reviewid" value="<%= oitem.FItemOne.Fidx %>">
<tr>
	<td width="100" align="center" bgcolor="#FFFFFF">ID</td>
	<td  bgcolor="#FFFFFF"> <%= oitem.FItemOne.Fidx %> </td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="#FFFFFF">상품 코드</td>
	<td bgcolor="#FFFFFF"><input type="text" name="itemid" value="<%= oitem.FItemOne.Fitemid %>">(없을시 공란)</td>
</tr>
<tr>
	<td width="80" align="center" bgcolor="#FFFFFF">메인등록일</td>
	<td bgcolor="#FFFFFF"><input type="text" size="10" name="vDate" value="<%= oitem.FItemOne.FViewDate %>" onClick="jsPopCal('vDate');" style="cursor:hand;"></td>
</tr>
<tr>
	<td width="80" align="center" bgcolor="#FFFFFF">사용여부</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="rd_nousing" value="Y" <% if oitem.FItemOne.Fisusing="" or oitem.FItemOne.Fisusing = "Y" then response.write "checked" %>>사용함
		<input type="radio" name="rd_nousing" value="N"<% if oitem.FItemOne.Fisusing = "N" then response.write "checked" %>>사용안함
	</td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="#FFFFFF">아이콘<br>(50x50)</td>
	<td bgcolor="#FFFFFF">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td>
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('icondiv','iconName','icon','50','50','false');"/>
				(<b><font color="red">50x50</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
				<input type="hidden" name="iconName" value="<%= oitem.FItemOne.Ficon %>">
				<div align="right" id="icondiv">
				<% if oitem.FItemOne.Ficon<>"" then %>
					<img src="<%= oitem.FItemOne.FIconUrl %>" width="50" height="50" onclick="showimage('<%= oitem.FItemOne.FIconUrl %>');" style="cursor:pointer">
				<% end if %>
				</div>

			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="#FFFFFF">당첨전 이미지<br>450x330</td>
	<td bgcolor="#FFFFFF">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td>
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imagediv','imageName','image','200','450','false');"/>
				(<b><font color="red">450x330</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
				<input type="hidden" name="imageName" value="<%= oitem.FItemOne.FImage %>">
				<div align="right" id="imagediv">
				<% if oitem.FItemOne.FImage<>"" then %>
					<img src="<%= oitem.FItemOne.FImageUrl %>" width="50" height="50" onclick="showimage('<%= oitem.FItemOne.FImageUrl %>');" style="cursor:pointer">
				<% end if %>
				</div>
			</td>
			<td></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="#FFFFFF">당첨이미지<br>450x330</td>
	<td bgcolor="#FFFFFF">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td>
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imageCdiv','imageCName','imageC','200','450','false');"/>
				(<b><font color="red">450x330</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
				<input type="hidden" name="imageCName" value="<%= oitem.FItemOne.Fimgconfirm %>">

				<div align="right" id="imageCdiv">
				<% if oitem.FItemOne.Fimgconfirm<>"" then %>
					<img src="<%= oitem.FItemOne.FImageConUrl %>" width="50" height="50" onclick="showimage('<%= oitem.FItemOne.FImageConUrl %>');" style="cursor:pointer">
				<% end if %>
				</div>
			</td>
			<td></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="#FFFFFF">다운로드이미지<br>966*?</td>
	<td bgcolor="#FFFFFF">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td>
				<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imageDdiv','imageDName','imageD','400','966','false');"/>
				(<b><font color="red">966*?</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
				<input type="hidden" name="imageDName" value="<%= oitem.FItemOne.FImgDown %>">

				<div align="right" id="imageDdiv">
				<% if oitem.FItemOne.FImgDown<>"" then %>
					<img src="<%= oitem.FItemOne.FImageDownUrl %>" width="50" height="50" onclick="showimage('<%= oitem.FItemOne.FImageDownUrl %>');" style="cursor:pointer">
				<% end if %>
				</div>
			</td>
			<td></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" align="center" bgcolor="#FFFFFF">
		<input type="button" value="저장" onClick="jsSubmit(frm_add)">
	</td>
</tr>
</form>
</table>
<form name="imginputfrm" method="post" action="">

<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
