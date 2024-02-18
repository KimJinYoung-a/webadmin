<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스
' History : 2009.04.07 서동석 생성
'			2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 600   'KB
dim ihashKey : ihashKey=getFimageHashKey(session("ssbctID")) '' 핑거스 이미지서버 권한 체크 관련 (도메인 다름)
dim lec_idx
lec_idx= requestCheckvar(request("lec_idx"),10)

public function GetImageSubFolderByItemid(byval lec_idx)
	GetImageSubFolderByItemid = "0" + CStr(Clng(lec_idx\10000))
end function

dim FLec_idx,Flec_title
dim Fbasicimg,Ficon1,Ficon2,Flistimg,Fsmallimg
dim Fmainimg,Fstoryimg
dim Foblong_img1,Foblong_img2,Foblong_img3,Foblong_img4
dim FBestLecimg, FNewLecimg
dim Faddimg1,Faddimg2,Faddimg3,Faddimg4,Faddimg5
dim FaddContents1,FaddContents2,FaddContents3,FaddContents4,FaddContents5
dim Fmorollingimg1, Fmorollingimg2, Fmorollingimg3

dim sql
sql=	" select idx, lec_title, basicimg,icon1,listimg,icon2,smallimg,mainimg,storyimg,addimg1,addimg2,addimg3,addimg4,addimg5" &_
			"	,oblongImg1,oblongImg2,oblongImg3,oblongImg4, morollingimg1, morollingimg2, morollingimg3 " &_
			"	,addcontents1,addcontents2,addcontents3,addcontents4,addcontents5 " &_
			" from db_academy.dbo.tbl_lec_item" &_
			" where idx='" & CStr(lec_idx) & "'"
rsACADEMYget.open sql,dbACADEMYget,1

if not rsACADEMYget.eof or not rsACADEMYget.bof then

	FLec_idx	= rsACADEMYget("idx")
	Flec_title	=	rsACADEMYget("lec_title")

	'// 정사각 이미지
	if Not(rsACADEMYget("basicimg")="" or isNull(rsACADEMYget("basicimg"))) then		Fbasicimg		= imgFingers & "/lectureitem/basic/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("basicimg")
	if Not(rsACADEMYget("icon1")="" or isNull(rsACADEMYget("icon1"))) then				Ficon1			= imgFingers & "/lectureitem/icon1/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("icon1")
	if Not(rsACADEMYget("icon2")="" or isNull(rsACADEMYget("icon2"))) then				Ficon2			= imgFingers & "/lectureitem/icon2/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("icon2")
	if Not(rsACADEMYget("listimg")="" or isNull(rsACADEMYget("listimg"))) then			Flistimg		= imgFingers & "/lectureitem/list/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("listimg")
	if Not(rsACADEMYget("smallimg")="" or isNull(rsACADEMYget("smallimg"))) then		Fsmallimg		= imgFingers & "/lectureitem/small/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("smallimg")

	'// 직사각(3:2) 이미지
	if Not(rsACADEMYget("oblongImg1")="" or isNull(rsACADEMYget("oblongImg1"))) then	Foblong_img1	= imgFingers & "/lectureitem/obl1/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("oblongImg1")
	if Not(rsACADEMYget("oblongImg2")="" or isNull(rsACADEMYget("oblongImg2"))) then	Foblong_img2	= imgFingers & "/lectureitem/obl2/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("oblongImg2")
	if Not(rsACADEMYget("oblongImg3")="" or isNull(rsACADEMYget("oblongImg3"))) then	Foblong_img3	= imgFingers & "/lectureitem/obl3/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("oblongImg3")
	if Not(rsACADEMYget("oblongImg4")="" or isNull(rsACADEMYget("oblongImg4"))) then	Foblong_img4	= imgFingers & "/lectureitem/obl4/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("oblongImg4")

	if Not(rsACADEMYget("mainimg")="" or isNull(rsACADEMYget("mainimg"))) then			Fmainimg		= imgFingers & "/lectureitem/main/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("mainimg")
	if Not(rsACADEMYget("storyimg")="" or isNull(rsACADEMYget("storyimg"))) then		Fstoryimg		= imgFingers & "/lectureitem/story1/" + GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("storyimg")
	if Not(rsACADEMYget("addimg1")="" or isNull(rsACADEMYget("addimg1"))) then			Faddimg1		= imgFingers & "/lectureitem/add1/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("addimg1")
	if Not(rsACADEMYget("addimg2")="" or isNull(rsACADEMYget("addimg2"))) then			Faddimg2		= imgFingers & "/lectureitem/add2/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("addimg2")
	if Not(rsACADEMYget("addimg3")="" or isNull(rsACADEMYget("addimg3"))) then			Faddimg3		= imgFingers & "/lectureitem/add3/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("addimg3")
	if Not(rsACADEMYget("addimg4")="" or isNull(rsACADEMYget("addimg4"))) then			Faddimg4		= imgFingers & "/lectureitem/add4/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("addimg4")
	if Not(rsACADEMYget("addimg5")="" or isNull(rsACADEMYget("addimg5"))) then			Faddimg5		= imgFingers & "/lectureitem/add5/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("addimg5")

	if rsACADEMYget("basicimg") <> "" then
		FBestLecimg 	= imgFingers & "/lectureitem/Index/"	+ GetImageSubFolderByItemid(lec_idx) + "/BL_" + mid(rsACADEMYget("basicimg"),5,len(rsACADEMYget("basicimg")))
		FNewLecimg 		= imgFingers & "/lectureitem/Index/"	+ GetImageSubFolderByItemid(lec_idx) + "/NL_" + mid(rsACADEMYget("basicimg"),5,len(rsACADEMYget("basicimg")))
	end if
	FaddContents1=db2html(rsACADEMYget("addcontents1"))
	FaddContents2=db2html(rsACADEMYget("addcontents2"))
	FaddContents3=db2html(rsACADEMYget("addcontents3"))
	FaddContents4=db2html(rsACADEMYget("addcontents4"))
	FaddContents5=db2html(rsACADEMYget("addcontents5"))

	'2016-05-24 모바일 롤링이미지1,2,3 유태욱 추가
	if Not(rsACADEMYget("morollingimg1")="" or isNull(rsACADEMYget("morollingimg1"))) then	Fmorollingimg1 = imgFingers & "/lectureitem/morolling1/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("morollingimg1")
	if Not(rsACADEMYget("morollingimg2")="" or isNull(rsACADEMYget("morollingimg2"))) then	Fmorollingimg2 = imgFingers & "/lectureitem/morolling2/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("morollingimg2")
	if Not(rsACADEMYget("morollingimg3")="" or isNull(rsACADEMYget("morollingimg3"))) then	Fmorollingimg3 = imgFingers & "/lectureitem/morolling3/" 	+ GetImageSubFolderByItemid(lec_idx) + "/" + rsACADEMYget("morollingimg3")
%>

<style>
.img_a {border:1px solid #BABABA}
}
</style>

<form name="lecfrm" method="post" action="<%=UploadImgFingers%>/linkweb/doFingerLecture_imgreg.asp?<%=ihashKey%>" enctype="multipart/form-data" style="margin:0px;">
<input type="hidden" name="mode" value="modi">
<input type="hidden" name="lec_idx" value="<%= lec_idx %>">

<table width="800" border="0" align="center"  class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#DDDDFF">
	<td width="150">강좌번호</td>
	<td width="650" bgcolor="#FFFFFF" align="left" colspan="2"><%= Flec_idx %>

	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="150">강좌명</td>
	<td width="650" bgcolor="#FFFFFF" align="left" colspan="2"><%= Flec_Title %>

	</td>
</tr>
<tr align="center">
	<td width="150" bgcolor="#DDDDFF" rowspan="2">정사각 이미지</td>
	<td width="50" bgcolor="#EFEFFF">기본</td>
	<td width="600" bgcolor="#FFFFFF" align="left">
		<input type="file" name="basicimg">(<font color="red">필수입력</font>, size : 400X400)<br>
		<% if Fbasicimg<>"" then %><img src="<%= Fbasicimg %>" class="img_a" border="0"><% end if %></td>
</tr>
<tr align="center">
	<td width="50" bgcolor="#EFEFFF">아이콘</td>
	<td width="600" bgcolor="#FFFFFF" align="left">
		<% if Ficon1<>"" then %><img src="<%= Ficon1 %>" class="img_a" border="0">&nbsp;<% end if %>
		<% if Flistimg<>"" then %><img src="<%= Flistimg %>" class="img_a" border="0">&nbsp;<% end if %>
		<% if Ficon2<>"" then %><img src="<%= Ficon2 %>" class="img_a" border="0">&nbsp;<% end if %>
		<% if Fsmallimg<>"" then %><img src="<%= Fsmallimg %>" class="img_a" border="0"><% end if %>
	</td>
</tr>
<tr align="center">
	<td bgcolor="#DDDDFF" >상품페이지<br>직사각(3:2)<br>이미지</td>
	<td bgcolor="#EFEFFF"></td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="file" name="storyimg">(<font color="red"></font>, size : 480X320)<br>
		<% if Fstoryimg<>"" then %><img src="<%= Fstoryimg %>" class="img_a" border="0"><% end if %>
	</td>
</tr>
		
<tr align="center">
	<td bgcolor="#DDDDFF" rowspan="2">직사각(3:2)<br>이미지</td>
	<td bgcolor="#EFEFFF">기본</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="file" name="oblongImg1">(<font color="red">필수입력</font>, size : 480X320)<br>
		<% if Foblong_img1<>"" then %><img src="<%= Foblong_img1 %>" class="img_a" border="0"><% end if %></td>
</tr>
<tr align="center">
	<td bgcolor="#EFEFFF">아이콘</td>
	<td bgcolor="#FFFFFF" align="left">
		<% if Foblong_img2<>"" then %><img src="<%= Foblong_img2 %>" class="img_a" border="0">&nbsp;<% end if %>
		<% if Foblong_img3<>"" then %><img src="<%= Foblong_img3 %>" class="img_a" border="0">&nbsp;<% end if %>
		<% if Foblong_img4<>"" then %><img src="<%= Foblong_img4 %>" class="img_a" border="0"><% end if %>
	</td>
</tr>
<!-------- // 2009리뉴얼 이후로는 사용안함 // --------->
<tr>
	<td bgcolor="#D0D0E0" colspan="3" align="center" onClick="swOldImgView()" style="cursor:pointer;font-size:10px;padding:0px;">
		<span id="dis00">UNUSED IMAGE VIEW ▼</span>
		<script language="javascript">
			function swOldImgView() {
				var dsf = document.all;
				if(dsf.dis01.style.display=="") {
					dsf.dis00.innerHTML = "UNUSED IMAGE VIEW ▼";
					dsf.dis01.style.display="none";
				} else {
					dsf.dis00.innerHTML = "UNUSED IMAGE VIEW ▲";
					dsf.dis01.style.display="";
				}
			}
		</script>
	</td>
</tr>
<tr id="dis01" style="display:none">
	<td colspan="3" style="padding:0px">
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1">
		<tr align="center" bgcolor="#D0D0D0">
			<td width="92">제품설명이미지</td>
			<td bgcolor="#F0F0F0" align="left" colspan="2">
				<input type="file" name="mainimg">(<font color="red">선택입력</font>, size : 300x250)<br>
				<% if Fmainimg<>"" then %><img src="<%= Fmainimg %>" class="img_a" border="0"><% end if %>
			</td>
		</tr>
		<!-- Story :2010-11 리뉴얼 이후 상품페이지 직사각 480x320 용으로 사용.. -->
		<!--
		<tr align="center" bgcolor="#D0D0D0">
			<td>Story 이미지</td>
			<td bgcolor="#F0F0F0" align="left" colspan="2">
				<input type="file" name="storyimg">(<font color="red">선택입력</font>, size : 150x110)<br>
				<% if Fstoryimg<>"" then %><img src="<%= Fstoryimg %>" class="img_a" border="0"><% end if %>
			</td>
		</tr>
		-->
		<tr align="center" bgcolor="#D0D0D0">
			<td rowspan="4">Main 이미지</td>
			<td bgcolor="#D0D0D0" align="left" colspan="2">
				인기강좌 이미지(<font color="red">선택입력</font>, size : 120x80, 기본이미지와 동일 확장자명)
			</td>
		</tr>
		<tr align="center" bgcolor="#D0D0D0">
			<td bgcolor="#F0F0F0" align="left" colspan="2">
				<input type="file" name="bestimg"><br>
				<% if FBestLecimg<>"" then %><img src="<%=FBestLecimg%>" class="img_a" border="0"><% end if %>
			</td>
		</tr>
		<tr align="center" bgcolor="#D0D0D0">
			<td bgcolor="#D0D0D0" align="left" colspan="2">
				새로운강좌 이미지(<font color="red">선택입력</font>, size : 50x70, 기본이미지와 동일 확장자명)
			</td>
		</tr>
		<tr align="center" bgcolor="#D0D0D0">
			<td bgcolor="#F0F0F0" align="left" colspan="2">
				<input type="file" name="newimg"><br>
				<% if FNewLecimg<>"" then %><img src="<%= FNewLecimg %>" class="img_a" border="0"><% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<!-------- // 2009리뉴얼 이후로는 사용안함 // --------->
<tr align="center" bgcolor="#DDDDFF">
	<td>추가 이미지1(상세1)</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="addimg1">(선택입력, size : 가로1000px,500kb이하)<br>
		<% if Faddimg1<>"" then %><img src="<%= Faddimg1 %>" class="img_a" border="0"><br><% end if %>
		설명: <textarea name="addcontents1" cols="70" rows="5"><%= Faddcontents1 %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td>추가 이미지2(상세2)</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="addimg2">(선택입력, size : 가로1000px,500kb이하)><br>
		<% if Faddimg2<>"" then %><img src="<%= Faddimg2 %>" class="img_a" border="0"><br><% end if %>
		설명: <textarea name="addcontents2" cols="70" rows="5"><%= Faddcontents2 %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td>추가 이미지3(상세3)</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="addimg3">(선택입력, size : 가로1000px,500kb이하)<br>
		<% if Faddimg3<>"" then %><img src="<%= Faddimg3 %>" class="img_a" border="0"><br><% end if %>
		설명: <textarea name="addcontents3" cols="70" rows="5"><%= Faddcontents3 %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td>추가 이미지4(상세4)</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="addimg4">(선택입력, size : 가로1000px,500kb이하)<br>
		<% if Faddimg4<>"" then %><img src="<%= Faddimg4 %>" class="img_a" border="0"><br><% end if %>
		설명: <textarea name="addcontents4" cols="70" rows="5"><%= Faddcontents4 %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td>추가 이미지5(상세5)</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="addimg5">(선택입력, size : 가로1000px,500kb이하)<br>
		<% if Faddimg5<>"" then %><img src="<%= Faddimg5 %>" class="img_a" border="0"><br><% end if %>
		설명: <textarea name="addcontents5" cols="70" rows="5"><%= Faddcontents5 %></textarea>
	</td>
</tr>
</table>
<% rsACADEMYget.close %>
<%
	Dim cImg, k, vArr, j, txtBuf
	set cImg = new CItemAddImage
	cImg.FRectItemID = lec_idx
	vArr = cImg.GetAddImageList
%>
<table width="800" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
<%
If isArray(vArr) Then
	If vArr(3,UBound(vArr,2)) > 0 Then
	For k = 1 To vArr(3,UBound(vArr,2))
%>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">추가 이미지<%= (k+5) %>(상세<%= (k+5) %>)</td>
	<td bgcolor="#FFFFFF">
	<%
	If cImg.IsImgExist(vArr,k) Then
		For j = 0 To UBound(vArr,2)
			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
				Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:block;""><img src=""" & imgFingers & "/lectureitem/contentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """></div>"
				Exit For
			End If
		Next
	Else
		Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:none;""></div>"
	End If
	%>
	  <input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40, <%= (k-1) %>);" class="text" size="40"> (선택입력, size : 가로1000px,500kb이하)
	  <br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/>
	  <%
	  txtBuf=""
	  For j = 0 To UBound(vArr,2)
		If CStr(vArr(3,j)) = CStr(k) Then
			txtBuf = vArr(5,j)
			Exit For
		End If
	  Next
	  %>
	  <textarea name="addimgtext" cols="70" rows="5"><%=txtBuf%></textarea>
	  <input type="hidden" name="addimggubun" value="<%= (k) %>">
	  <input type="hidden" name="addimgdel" value="">
	</td>
</tr>
<%
	Next
	End IF
End IF
%>
</table>
<%	set cImg = nothing %>
<script type="text/javascript">
<!--
//상품설명이미지추가
function InsertImageUp() {
	var f = document.all;
	var rowLen = f.imgIn.rows.length;

	if(rowLen > 9){
		alert("이미지는 최대 15개까지 가능합니다.");
		return;
	}
	//rowLen=1;
	var i = rowLen;
	var r  = f.imgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var c1 = r.insertCell(1);

	c0.style.textAlign = 'center';
	c1.style.textAlign = 'left';
	c0.style.height = '30';
	c0.style.width = '150';
	c1.style.width = '650';
	c0.style.background = '#DDDDFF';
	c0.innerHTML = '추가 이미지' + (rowLen+5) + '(상세' + (rowLen+5) + ')';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40, '+parseInt(rowLen-1)+');"> ';
	c1.innerHTML += ' (선택입력, size : 가로1000px,500kb이하)';
	c1.innerHTML += '<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>';
	c1.innerHTML += '<input type="hidden" name="addimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addimgdel" value="">';
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
//-->
</script>
<table width="800" border="0" align="center"  class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td colspan="3" bgcolor="#FFFFFF">
	<input type="button" value="상세이미지추가" class="button" onClick="InsertImageUp()">
		  <font color="red">* 업로드가 된 이미지가 제대로 안나오면 새로고침(CTRL + F5(콘트롤 F5 버튼))을 해주세요.</font>
	</td>
</tr>
<!--모바일 롤링 이미지1,2,3 2016-05-24 유태욱 추가-->
<tr align="center" bgcolor="#DDDDFF">
	<td>모바일 상세롤링1 & <br>리스트 이미지</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="morollingimg1">(size : 1000x667)<br>
		<% if Fmorollingimg1<>"" then %><img src="<%= Fmorollingimg1 %>" width="500" class="img_a" border="0"><br><% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td>모바일 상세롤링2</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="morollingimg2">(size : 1000x667)<br>
		<% if Fmorollingimg2<>"" then %><img src="<%= Fmorollingimg2 %>" width="500" class="img_a" border="0"><br><% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td>모바일 상세롤링3</td>
	<td bgcolor="#FFFFFF" align="left" colspan="2">
		<input type="file" name="morollingimg3">(size : 1000x667)<br>
		<% if Fmorollingimg3<>"" then %><img src="<%= Fmorollingimg3 %>" width="500" class="img_a" border="0"><br><% end if %>
	</td>
</tr>

<tr>
	<td colspan="3" bgcolor="#FFFFFF" align="center"><input type="submit" value="저장" ></td>
</tr>
</table>

</form>

<%
end if


%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->