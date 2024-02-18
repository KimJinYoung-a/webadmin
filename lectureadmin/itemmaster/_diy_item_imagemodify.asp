<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%
'CONST CBASIC_IMG_MAXSIZE = 180   'KB
'CONST CMAIN_IMG_MAXSIZE = 500   'KB

CONST CBASIC_IMG_MAXSIZE = 600   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim itemid, oitem

itemid = requestCheckvar(request("itemid"),10)
'session("ssBctID")


'==============================================================================
if (itemid = "") then
    itemid = -1
end if


'==============================================================================
set oitem = new CItem

'TODO : 업체페이지로 이동시 위에 include 부분과, 아래 코맨트를 바꿔준다.
oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemID = itemid
if (oitem.FRectMakerId<>"") then
    oitem.GetOneItem
end if

if (oitem.FResultCount < 1) then
    response.write "<script>alert('잘못된 접속입니다.');</script>"
    dbACADEMYget.close()	:	response.End
end if



dim oitemAddImage
set oitemAddImage = new CItemAddImage
oitemAddImage.FRectItemID = itemid

if oitem.FResultCount>0 then
    ''상품 추가이미지 접수.
    oitemAddImage.GetOneItemAddImageList
end if

'==============================================================================
dim i
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

    if (itemreg.basic.value == "del") {
        alert("기본이미지는 필수입니다.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg') != true) {
                return;
            }
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif') != true) {
            return;
        }
    }

//    if (itemreg.imgadd4.value != "") {
//        if (CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgadd5.value != "") {
//        if (CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif') != true) {
//            return;
//        }
//    }
//
//    if (itemreg.imgmain.value != "") {
//        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif') != true) {
//            return;
//        }
//    }

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
function ClearImage(img) {
    var e = eval("itemreg." + img);
    // TODO : 아래방식이 깔끔하지만 에러가 난다. ㅡㅡ;;
    // e.select();
    // document.execCommand('Delete');

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);\" size='40'>";
    }


    e = eval("document.all.div" + img);
    e.style.display = "none";

    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "del";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "del";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "del";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "del";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "del";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "del";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "del";
    }
}

function ClearImage2(img,fsize,wd,ht,num) {
	var imgcnt = $('input[name="addimgname"]').length;
    img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +", "+num+");CheckImageSize(this);\" class='text' size='"+ fsize +"'>";
	$("#divaddimgname"+(num+1)+"").hide();
	
	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "del";
    }else{
    	document.itemreg.addimgdel.value = "del";
    }
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

function CheckImage(img, filesize, imagewidth, imageheight, extname)
{
    var preview;
    var e;
    var ext;
    var filename;

    e = eval("itemreg." + img);
    filename = e.value;

    e = eval("itemreg." + img);
    if (e.value == "") { return false; }

    if (CheckExtension(filename, extname) != true) {
        alert("이미지화일은 다음의 화일만 사용하세요.[" + extname + "]");
        ClearImage(img);
        return false;
    }

    if (img == "imgbasic") {
        e = eval("itemreg.basic");
        e.value = "";
    } else if (img == "imgadd1") {
        e = eval("itemreg.add1");
        e.value = "";
    } else if (img == "imgadd2") {
        e = eval("itemreg.add2");
        e.value = "";
    } else if (img == "imgadd3") {
        e = eval("itemreg.add3");
        e.value = "";
    } else if (img == "imgadd4") {
        e = eval("itemreg.add4");
        e.value = "";
    } else if (img == "imgadd5") {
        e = eval("itemreg.add5");
        e.value = "";
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "";
    }

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

//상품설명이미지추가
function InsertImageUp() {
	var f = document.all;
	var rowLen = f.imgIn.rows.length;

	if(rowLen > 14){
		alert("이미지는 최대 15개까지 가능합니다.");
		return;
	}
	
	var i = rowLen;
	var r  = f.imgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var c1 = r.insertCell(1);

	r.style.textAlign = 'left';
	c0.style.height = '30';
	c0.style.width = '15%';
	c0.style.background = '#DDDDFF';
	c0.innerHTML = '상품상세이미지 #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40, '+parseInt(rowLen-1)+');CheckImageSize(this);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 1000, 667, '+parseInt(rowLen-1)+')"> (선택,1000X667, Max 600KB,jpg,gif)';
	c1.innerHTML += '<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>';
	c1.innerHTML += '<input type="hidden" name="addimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addimgdel" value="">';
}

function CheckImageSize(obj) {
	var MaxSize=600;
	if((obj.files[0].size/1024) > MaxSize){
		alert("이미지는 600kb 까지 올리실 수 있습니다. (" + ((obj.files[0].size/1024)-MaxSize).toFixed(2) + "kb 초과)" );
		obj.value="";
		return;
	}
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


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>이미지정보
          <br>- 텐바이텐에서 이미지를 등록할 경우 따로 입력하지 마시기 바랍니다.

          <br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
          <br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
          <br>- <font color=red>포토샾에서 Save For Web으로, Optimize체크, 압축율 80%이하</font>로 만드신 후 올려주시기 바랍니다.
          <br><br>이미지 수정후 <font color=red>CTRL + F5 (콘트롤 F5 버튼)</font></a> 누르셔야 수정된 이미지를 확인하실 수 있습니다
          <br><br><input type="button" value=" 새로고침 " onClick="location.reload();"> &nbsp;&nbsp; <input type="button" value=" 닫 기 " onClick="CloseWindow()"><br>&nbsp;


        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->
<% if (TRUE) or (session("ssBCTid")="fingertest01") then %>
<form name="itemreg" method="post" action="<%= uploadImgUrl  %>/linkweb/academy/items/DIYItemImageModify.asp" enctype="MULTIPART/FORM-DATA">
<% else %>
<form name="itemreg" method="post" action="<%= UploadImgFingers  %>/linkweb/items/DIYItemImageModify.asp" enctype="MULTIPART/FORM-DATA">
<% end if %>
<input type="hidden" name="itemid" value="<%= itemid %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
	  <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg');CheckImageSize(this);" size="40"> (<font color=red>필수</font>,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic">

<% if (oitem.FOneItem.FbasicImage <> "") then %>
      <div id="divimgbasic" style="display:block;">
      <table width="100%" height="400" >
        <tr>
          <td><img id="imgbasic_img" src="<%= oitem.FOneItem.FbasicImage %>"></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgbasic" style="display:none;">
      <table width="100%" height="400" >
        <tr>
          <td><img id="imgbasic_img" src=""></td>
        </tr>
      </table>
      </div>
<% end if %>
    </td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">아이콘이미지<br>(자동생성) :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (oitem.FOneItem.Ficon1image <> "") then %>
      <img src="<%= oitem.FOneItem.Ficon1image %>" width="360" >
<% end if %>
<% if (oitem.FOneItem.Ficon2image <> "") then %>
      <img src="<%= oitem.FOneItem.Ficon2image %>" width="195">
<% end if %>
<% if (oitem.FOneItem.Flistimage120 <> "") then %>
      <img src="<%= oitem.FOneItem.Flistimage120 %>" width="150">
<% end if %>
<% if (oitem.FOneItem.Flistimage <> "") then %>
      <img src="<%= oitem.FOneItem.Flistimage %>" width="75">
<% end if %>

<% if (oitem.FOneItem.Fsmallimage <> "") then %>
      <img src="<%= oitem.FOneItem.Fsmallimage %>" width="60">
<% end if %>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1">
<% if (oitemAddImage.GetImageAddByIdx(0,1) <> "") then %>
      <div id="divimgadd1" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd1_img" src="<%= oitemAddImage.GetImageAddByIdx(0,1) %>"></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd1" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd1_img" src=""></td>
        </tr>
      </table>
      </div>
<% end if %>
    </td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2">
<% if (oitemAddImage.GetImageAddByIdx(0,2) <> "") then %>
      <div id="divimgadd2" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd2_img" src="<%= oitemAddImage.GetImageAddByIdx(0,2) %>"></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgadd2" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd2_img" src=""></td>
        </tr>
        
      </table>
      </div>
<% end if %>

  	</td>
  </tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');CheckImageSize(this);" size="40"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3">
<% if (oitemAddImage.GetImageAddByIdx(0,3) <> "") then %>
      <div id="divimgadd3" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd3_img" src="<%= oitemAddImage.GetImageAddByIdx(0,3) %>"></td>
        </tr>
       
      </table>
      </div>
<% else %>
      <div id="divimgadd3" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd3_img" src=""></td>
        </tr>
        
      </table>
      </div>
<% end if %>

  	</td>
  </tr>

<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--   	  <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> (선택,1000X667,jpg,gif) -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4"> -->
<!-- <% if (oitemAddImage.GetImageAddByIdx(0,4) <> "") then %> -->
<!--       <div id="divimgadd4" style="display:block;"> -->
<!--       <table width="100%" height="400" class="a"> -->
<!--         <tr> -->
<!--           <td colspan="2"><img id="imgadd4_img" src="<%= oitemAddImage.GetImageAddByIdx(0,4) %>"></td> -->
<!--         </tr> -->
<!--          -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgadd4" style="display:none;"> -->
<!--       <table width="100%" height="400" class="a"> -->
<!--         <tr> -->
<!--           <td colspan="2"><img id="imgadd4_img" src=""></td> -->
<!--         </tr> -->
<!--          -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--  -->
<!--   	</td> -->
<!--   </tr> -->
<!--  -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지5 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--   	  <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif');" size="40"> (선택,1000X667,jpg,gif) -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5"> -->
<!-- <% if (oitemAddImage.GetImageAddByIdx(0,5) <> "") then %> -->
<!--       <div id="divimgadd5" style="display:block;"> -->
<!--       <table width="100%" height="400" class="a"> -->
<!--         <tr> -->
<!--           <td colspan="2"><img id="imgadd5_img" src="<%= oitemAddImage.GetImageAddByIdx(0,5) %>"></td> -->
<!--         </tr> -->
<!--          -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgadd5" style="display:none;"> -->
<!--       <table width="100%" height="400" class="a"> -->
<!--         <tr> -->
<!--           <td colspan="2"><img id="imgadd5_img" src=""></td> -->
<!--         </tr> -->
<!--          -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--  -->
<!--   	</td> -->
<!--   </tr> -->
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--   	  <input type="file" name="imgmain" onchange="CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');" size="40"> (선택,600X2000,<%= CMAIN_IMG_MAXSIZE %>KB,jpg, gif) -->
<!--       <input type="button" value="이미지지우기" onClick="ClearImage('imgmain')"><input type="hidden" name="main"> -->
<!-- <% if (oitem.FOneItem.Fmainimage <> "") then %> -->
<!--       <div id="divimgmain" style="display:block;"> -->
<!--       <table width="100%" height="400" > -->
<!--         <tr> -->
<!--           <td><img id="imgmain_img" src="<%= oitem.FOneItem.Fmainimage %>"></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% else %> -->
<!--       <div id="divimgmain" style="display:none;"> -->
<!--       <table width="100%" height="400" > -->
<!--         <tr> -->
<!--           <td><img id="imgmain_img" src=""></td> -->
<!--         </tr> -->
<!--       </table> -->
<!--       </div> -->
<!-- <% end if %> -->
<!--   	</td> -->
<!--   </tr> -->
</table>
<%
	Dim cImg, k, vArr, j, txtBuf
	set cImg = new CItemAddImage
	cImg.FRectItemID = itemid
	vArr = cImg.GetAddImageList
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
	<% If isArray(vArr) Then
			If vArr(3,UBound(vArr,2)) > 0 Then
			For k = 1 To vArr(3,UBound(vArr,2))
	%>
			  <tr align="left">
			  	<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #<%= (k) %> :</td>
			  	<td bgcolor="#FFFFFF">
		  		<%
		  		If cImg.IsImgExist(vArr,k) Then
		    		For j = 0 To UBound(vArr,2)
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:block;""><img src=""" & imgFingers & "/diyItem/contentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ height=""250""></div>"
							Exit For
		    			End If
		    		Next
				Else
					Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:none;""></div>"
				End If
				%>
			      <input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40, <%= (k-1) %>);CheckImageSize(this);" class="text" size="40">
			      <input type="button" value="#<%= (k) %> 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname<%=CHKIIF(vArr(3,UBound(vArr,2))=1,"","["&(k-1)&"]")%>,40, 1000, 667, <%= (k-1) %>)"> (선택,1000X667, Max 800KB,jpg,gif)
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
		Else
	%>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #1 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname1" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,0);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#1 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 1000, 667, 0)"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="1">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #2 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname2" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,1);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#2 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 1000, 667, 1)"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="2">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #3 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname3" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,2);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#3 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 1000, 667, 2)"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="3">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #4 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname4" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,3);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#4 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[3],40, 1000, 667, 3)"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="4">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #5 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname5" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,4);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#5 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[4],40, 1000, 667, 4)"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="5">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품상세이미지 #6 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname6" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,5);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#6 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[5],40, 1000, 667, 5)"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="6">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">상품설명이미지 #7 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname7" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 667, 'jpg,gif',40,6);CheckImageSize(this);" class="text" size="40">
				<input type="button" value="#7 이미지지우기" class="button" onClick="ClearImage2(this.form.addimgname[6],40, 1000, 667, 6)"> (선택,1000X667,Max <%= CBASIC_IMG_MAXSIZE %>KB,jpg,gif)
				<br/><span style="color:red;font-size:15px"><strong>※이미지 등록 없이 설명만 올릴 수 없습니다.※</strong></span><br/><textarea name="addimgtext" cols="70" rows="5"></textarea>
				<input type="hidden" name="addimggubun" value="7">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
	<%
	   End IF %>
</table>
<%	set cImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr align="left">
  	<td bgcolor="#FFFFFF" height="30">
      <input type="button" value="상품상세이미지추가" class="button" onClick="InsertImageUp()">
      <font color="red">* 업로드가 된 이미지가 제대로 안나오면 새로고침(CTRL + F5(콘트롤 F5 버튼))을 해주세요.</font>
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
set oitem = Nothing
set oitemAddImage = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->