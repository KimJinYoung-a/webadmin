<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.04.15 한용민 생성
'	Description : 감성엽서
'#######################################################
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
<!-- #include virtual="/lib/classes/sens/image_commentcls.asp"-->
<%
dim reviewid, mode
	reviewid = request("reviewid")
	mode = request("mode")

if mode="add" then reviewid=0

dim oitem
	set oitem = new CItemImage

oitem.FRectReviewID = reviewid
oitem.GetOneItemImage

dim yyyy1,mm1,dd1,viewdate,viewdatearr
	viewdate = FormatDateTime(oitem.FItemOne.Fviewdate,2)
	viewdatearr = split(viewdate,"-")

if viewdate = "" then
	if (yyyy1="") then yyyy1 = Cstr(Year(now()))
	if (mm1="") then mm1 = Cstr(Month(now()))
	if (dd1="") then dd1 = Cstr(day(now()))
else
	yyyy1 = viewdatearr(0)
	mm1 = viewdatearr(1)
	dd1 = viewdatearr(2)
end if
%>

<script language='javascript'>

function DelImage(src, imgname, hid){
	var imgcomp;
	imgcomp = eval("document." + imgname);
	if (hid.checked){
		imgcomp.src = '/images/space.gif';
	}
}

function TnEditFinger(frm){

	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		frm.submit();
	}
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes');
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
	document.imginputfrm.action='postcard_img_input.asp';
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
	//document.imginputfrm.action='http://upload.10x10.co.kr/linkweb/sens/sens_postcard_img_del.asp';
	//document.imginputfrm.submit();
}

document.domain='10x10.co.kr';
</script>

<table width="700" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#B2B2B2">
	<form name="frm_add" method="post" action="proc_image_comment.asp">
	<input type="hidden" name="mode" value="<%= mode %>">
	<input type="hidden" name="reviewid" value="<%= oitem.FItemOne.Fidx %>">
	<tr>
		<td align="center" bgcolor="#FFFFFF">ID</td>
		<td  bgcolor="#FFFFFF"> <%= oitem.FItemOne.Fidx %> </td>
	</tr>
	<tr>
		<td align="center" bgcolor="#FFFFFF">상품 코드</td>
		<td bgcolor="#FFFFFF"><input type="text" name="itemid" value="<%= oitem.FItemOne.Fitemid %>"> (없을경우 공백으로 두세요)</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#FFFFFF">노출시작일</td>
		<td bgcolor="#FFFFFF"><% DrawOneDateBox yyyy1,mm1,dd1 %></td>
	</tr>
	<tr>
		<td align="center" bgcolor="#FFFFFF">사용여부</td>
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
					<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('icondiv','iconName','icon','200','50','false');"/>
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
		<td width="100" align="center" bgcolor="#FFFFFF">당첨전 이미지<br>459x358</td>
		<td bgcolor="#FFFFFF">
			<table border="0" cellspacing="0" cellpadding="0" class="a">
			<tr>
				<td>
					<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imagediv','imageName','image','300','460','false');"/>
					(<b><font color="red">459x358</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
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
		<td width="100" align="center" bgcolor="#FFFFFF">당첨이미지<br>(없을경우 위에이미지노출)<br>459x358</td>
		<td bgcolor="#FFFFFF">
			<table border="0" cellspacing="0" cellpadding="0" class="a">
			<tr>
				<td>
					<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imageCdiv','imageCName','imageC','300','460','false');"/>
					(<b><font color="red">459x358</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
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
		<td colspan="2" align="center" bgcolor="#FFFFFF">
			<input type="button" value="저장" onClick="TnEditFinger(frm_add)">
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