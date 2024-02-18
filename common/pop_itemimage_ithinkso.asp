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
<!-- #include virtual="/lib/classes/itemcls_v2.asp"-->
<!-- #include virtual="/lib/classes/itemoptioncls_v2.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 150   'KB
CONST CMAIN_IMG_MAXSIZE = 500   'KB

dim itemid, oitem
itemid = requestCheckVar(request("itemid"),10)
'session("ssBctID")


'==============================================================================
if (itemid = "") then
    itemid = -1
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetProductOne


'==============================================================================
dim i
%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type='text/javascript'>

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
            if (CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg') != true) {
                return;
            }
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage('imgadd1', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage('imgadd2', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage('imgadd3', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd4.value != "") {
        if (CheckImage('imgadd4', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif') != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage('imgadd5', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif') != true) {
            return;
        }
    }
    if (itemreg.imgadd6.value != "") {
        if (CheckImage('imgadd6', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif') != true) {
            return;
        }
    }
    

    if (itemreg.imgmain.value != "") {
        if (CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif') != true) {
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
function ClearImage(img) {
    var e = eval("itemreg." + img);
    // TODO : 아래방식이 깔끔하지만 에러가 난다. ㅡㅡ;;
    // e.select();
    // document.execCommand('Delete');

	if (img == "imgbasic") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');\" size='40'>";
    } else if (img == "imgmain") {
       e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');\" size='40'>";    
    } else if (img == "imgadd1") {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 910, 910, 'jpg,gif');\" size='40'>";
    } else {
        e.outerHTML="<input type='file' name='" + img + "' onchange=\"CheckImage('" + img + "', <%= CMAIN_IMG_MAXSIZE %>, 910, 910, 'jpg,gif');\" size='40'>";
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
    } else if (img == "imgadd6") {
        e = eval("itemreg.add6");
        e.value = "del";      
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "del";
    }
}

function ShowImage(img) {
    var e = eval("document.all.div" + img);
    e.style.display = "";

    var filename;
    e = eval("itemreg." + img);
    filename = e.value;

	eval("document.all." + img + "_img").src=filename;
    //document.getElementById(img).background=filename;
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

    ShowImage(img);

    // iframe 속에 이미지를 넣고, 사이즈/크기를 체크한다.
    document.imgpreview.document.getElementById("imgpreview").src = filename;
    preview = document.imgpreview.document.getElementById("imgpreview");

    if(preview.fileSize > (filesize * 1000)){
        alert("파일사이즈는 " + filesize + "Kbyte를 넘기실 수 없습니다.");
        ClearImage(img);
        return false;
    }
    	
    if(preview.width > (imagewidth)){
       alert("가로폭은 " + imagewidth + "픽셀을 넘기실 수 없습니다.");
       ClearImage(img);
       return false;
    }

    if(preview.height > (imageheight)){
       alert("세로폭은 " + imageheight + "픽셀을 넘기실 수 없습니다.");
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
    } else if (img == "imgadd6") {
        e = eval("itemreg.add6");
        e.value = "";     
    } else if (img == "imgmain") {
        e = eval("itemreg.main");
        e.value = "";
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


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>이미지정보
          <br>- 텐바이텐에서 이미지를 등록할 경우 따로 입력하지 마시기 바랍니다.

          <br>- 이미지는 <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> 까지 올리실 수 있습니다.
          <br>&nbsp;&nbsp;(이미지사이즈나 <font color=red>가로세로폭의 사이즈</font>를 규격에 넘지 않게 등록해주세요. 규격초과시 등록이 되지 않습니다.)
          <br>- <font color=red>포토샾에서 Save For Web</font>으로 만드신 후 올려주시기 바랍니다.
          <br><br>이미지 수정후 <font color=red>CTRL + F5 (콘트롤 F5 버튼)</font></a> 누르셔야 수정된 이미지를 확인하실 수 있습니다
          <br><br><input type="button" value=" 새로고침 " onClick="location.reload();"> &nbsp;&nbsp; <input type="button" value=" 닫 기 " onClick="CloseWindow()"><br>&nbsp;


        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <!--<form name="itemreg" method="post" action="http://testupload.10x10.co.kr/linkweb/itemmaster/doitemimagemodify_ithinkso.asp" enctype="MULTIPART/FORM-DATA">-->
  <form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/itemImageModify_ithinkso.asp" enctype="MULTIPART/FORM-DATA">
  <input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
  <tr>
	<td colspan="4" bgcolor="#F5EF80" height="25"> ithinkso 이미지 등록</td>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">기본이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <iframe name="imgpreview" src="iframe_imagepreview.asp" frameborder="0" width="0" height="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>
	  <input type="file" name="imgbasic" onchange="CheckImage('imgbasic', <%= CBASIC_IMG_MAXSIZE %>, 410, 410, 'jpg');" size="40"> (<font color=red>필수</font>,400X400,jpg)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgbasic')"><input type="hidden" name="basic">

<% if (oitem.FOneItem.Fimgbasic <> "") then %>
      <div id="divimgbasic" style="display:block;">
      <table width="100%" height="400" >
        <tr>
          <td><img id="imgbasic_img" src="<%= oitem.FOneItem.Fimgbasic %>"></td>
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
  	<td height="30" width="15%" bgcolor="#DDDDFF">아이콘이미지(자동생성) :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
<% if (oitem.FOneItem.Ficon1 <> "") then %>
      <img src="<%= oitem.FOneItem.Ficon1 %>" width="200" height="200">
<% end if %>
<% if (oitem.FOneItem.Ficon2 <> "") then %>
      <img src="<%= oitem.FOneItem.Ficon2 %>" >
<% end if %>
<% if (oitem.FOneItem.Flistimage120 <> "") then %>
      <img src="<%= oitem.FOneItem.Flistimage120 %>" width="120" height="120">
<% end if %>
<% if (oitem.FOneItem.Fimglist <> "") then %>
      <img src="<%= oitem.FOneItem.Fimglist %>" width="100" height="100">
<% end if %>
<% if (oitem.FOneItem.Fimgsmall <> "") then %>
      <img src="<%= oitem.FOneItem.Fimgsmall %>" width="50" height="50">
<% end if %>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지1 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd1" onchange="CheckImage('imgadd1', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif');" size="40"> (선택,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd1')"><input type="hidden" name="add1">      
<% if (oitem.GetImageAddByIndex(1) <> "") then %>
      <div id="divimgadd1" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd1_img" src="<%= oitem.GetImageAddByIndex(1) %>"></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30" >추가이미지설명1 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent1" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명1]" value="<%= oitem.GetImageContentByIndex(1) %>">
		  	</td>
		</tr>
		-->
      </table>
      </div>
<% else %>
      <div id="divimgadd1" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd1_img" src=""></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30" >추가이미지설명1 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent1" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명1]" value="<%= oitem.GetImageContentByIndex(1) %>">
		  	</td>
		</tr>
		-->
      </table>
      </div>
<% end if %>
    </td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지2 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd2" onchange="CheckImage('imgadd2', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif');" size="40"> (선택,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd2')"><input type="hidden" name="add2">
<% if (oitem.GetImageAddByIndex(2) <> "") then %>
      <div id="divimgadd2" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd2_img" src="<%= oitem.GetImageAddByIndex(2) %>"></td>
        </tr>
        <!--
        <tr align="left">
    	  	<td height="30" >추가이미지설명2 :</td>
    	  	<td bgcolor="#FFFFFF" >
    	      <input type="text" name="itemaddcontent2" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명2]" value="<%= oitem.GetImageContentByIndex(2) %>">
    	  	</td>
    	  </tr>
	    -->
      </table>
      </div>
<% else %>
      <div id="divimgadd2" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd2_img" src=""></td>
        </tr>
        <!--
        <tr align="left">
	  	<td height="30" >추가이미지설명2 :</td>
	  	<td bgcolor="#FFFFFF" >
	      <input type="text" name="itemaddcontent2" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명2]" value="<%= oitem.GetImageContentByIndex(2) %>">
	  	</td>
	  </tr>
	  -->
      </table>
      </div>
<% end if %>

  	</td>
  </tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지3 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd3" onchange="CheckImage('imgadd3', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif');" size="40"> (선택,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd3')"><input type="hidden" name="add3">
<% if (oitem.GetImageAddByIndex(3) <> "") then %>
      <div id="divimgadd3" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd3_img" src="<%= oitem.GetImageAddByIndex(3) %>"></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30" >추가이미지설명3 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent3" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명3]" value="<%= oitem.GetImageContentByIndex(3) %>">
		  	</td>
		  </tr>
		-->
      </table>
      </div>
<% else %>
      <div id="divimgadd3" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd3_img" src=""></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30" >추가이미지설명3 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent3" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명3]" value="<%= oitem.GetImageContentByIndex(3) %>">
		  	</td>
		  </tr>
		-->
      </table>
      </div>
<% end if %>

  	</td>
  </tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지4 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd4" onchange="CheckImage('imgadd4', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif');" size="40"> (선택,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd4')"><input type="hidden" name="add4">
<% if (oitem.GetImageAddByIndex(4) <> "") then %>
      <div id="divimgadd4" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd4_img" src="<%= oitem.GetImageAddByIndex(4) %>"></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30" >추가이미지설명4 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent4" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명4]" value="<%= oitem.GetImageContentByIndex(4) %>">
		  	</td>
		  </tr>
		-->
      </table>
      </div>
<% else %>
      <div id="divimgadd4" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd4_img" src=""></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30" >추가이미지설명4 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent4" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명4]" value="<%= oitem.GetImageContentByIndex(4) %>">
		  	</td>
		  </tr>
		  -->
      </table>
      </div>
<% end if %>

  	</td>
  </tr>

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지5 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd5" onchange="CheckImage('imgadd5', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif');" size="40"> (선택,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd5')"><input type="hidden" name="add5">
<% if (oitem.GetImageAddByIndex(5) <> "") then %>
      <div id="divimgadd5" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd5_img" src="<%= oitem.GetImageAddByIndex(5) %>"></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30">추가이미지설명5 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent5" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명5]" value="<%= oitem.GetImageContentByIndex(5) %>">
		  	</td>
		  </tr>
		 -->
      </table>
      </div>
<% else %>
      <div id="divimgadd5" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd5_img" src=""></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30">추가이미지설명5 :</td>
		  	<td bgcolor="#FFFFFF">
		      <input type="text" name="itemaddcontent5" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명5]" value="<%= oitem.GetImageContentByIndex(5) %>">
		  	</td>
		  </tr>
		-->
      </table>
      </div>
<% end if %>

  	</td>
  </tr>
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">추가이미지6 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgadd6" onchange="CheckImage('imgadd6', <%= CMAIN_IMG_MAXSIZE %>, 910, 1210, 'jpg,gif');" size="40"> (선택,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgadd6')"><input type="hidden" name="add6">
<% if (oitem.GetImageAddByIndex(6) <> "") then %>
      <div id="divimgadd6" style="display:block;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd6_img" src="<%= oitem.GetImageAddByIndex(6) %>"></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30">추가이미지설명5 :</td>
		  	<td bgcolor="#FFFFFF" >
		      <input type="text" name="itemaddcontent5" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명5]" value="<%= oitem.GetImageContentByIndex(5) %>">
		  	</td>
		  </tr>
		 -->
      </table>
      </div>
<% else %>
      <div id="divimgadd6" style="display:none;">
      <table width="100%" height="400" class="a">
        <tr>
          <td colspan="2"><img id="imgadd6_img" src=""></td>
        </tr>
        <!--
        <tr align="left">
		  	<td height="30">추가이미지설명5 :</td>
		  	<td bgcolor="#FFFFFF">
		      <input type="text" name="itemaddcontent5" size="50" maxlength="50" id="[off,off,off,off][추가이미지설명5]" value="<%= oitem.GetImageContentByIndex(5) %>">
		  	</td>
		  </tr>
		-->
      </table>
      </div>
<% end if %>

  	</td>
  </tr>
  <!-- <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">왼쪽메뉴이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgtitle" onchange="CheckImage('imgtitle', <%= CBASIC_IMG_MAXSIZE %>, 160, 160, 'jpg,gif');" size="40"> (선택,150X150,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgtitle')"><input type="hidden" name="title">
<% if (oitem.FOneItem.Fimgtitle <> "") then %>
      <div id="divimgtitle" style="display:block;">
      <table width="100%" height="150" >
        <tr>
          <td><img id="imgtitle_img" src="<%= oitem.FOneItem.Fimgtitle %>"></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgtitle" style="display:none;">
      <table width="100%" height="150" >
        <tr>
          <td><img id="imgtitle_img" src=""></td>
        </tr>
      </table>
      </div>
<% end if %>

  	</td>
  </tr>-->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">제품설명이미지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <input type="file" name="imgmain" onchange="CheckImage('imgmain', <%= CMAIN_IMG_MAXSIZE %>, 610, 2000, 'jpg,gif');" size="40"> (선택,600X2000,jpg,gif)
      <input type="button" value="이미지지우기" onClick="ClearImage('imgmain')"><input type="hidden" name="main">
<% if (oitem.FOneItem.Fimgmain <> "") then %>
      <div id="divimgmain" style="display:block;">
      <table width="100%" height="400" >
        <tr>
          <td><img id="imgmain_img" src="<%= oitem.FOneItem.Fimgmain %>"></td>
        </tr>
      </table>
      </div>
<% else %>
      <div id="divimgmain" style="display:none;">
      <table width="100%" height="400" >
        <tr>
          <td><img id="imgmain_img" src=""></td>
        </tr>
      </table>
      </div>
<% end if %>

  	</td>
  </tr>
  </form>
</table>


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
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->