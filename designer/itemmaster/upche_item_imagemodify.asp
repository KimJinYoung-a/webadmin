<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- include virtual="/lib/classes/itemoptioncls_v2.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
CONST CMAIN_IMG_MAXSIZE = 640   'KB

dim itemid, oitem

itemid = requestCheckVar(request("itemid"),20)
'session("ssBctID")


'==============================================================================
if (itemid = "") then
    itemid = -1
end if


'==============================================================================
set oitem = new CItem

'TODO : ��ü�������� �̵��� ���� include �κа�, �Ʒ� �ڸ�Ʈ�� �ٲ��ش�.
oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemID = itemid
oitem.GetOneItem

if (oitem.FResultCount < 1) then
    response.write "<script>alert('�߸��� �����Դϴ�.');</script>"
    dbget.close()	:	response.End
end if



dim oitemAddImage
set oitemAddImage = new CItemAddImage
oitemAddImage.FRectItemID = itemid

if oitem.FResultCount>0 then
    ''��ǰ �߰��̹��� ����.
    oitemAddImage.GetOneItemAddImageList
end if

'==============================================================================
dim i
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript'>

// ============================================================================
// �����ϱ�
function SubmitSave() {

    if (validate(itemreg)==false) {
        return;
    }

    if (itemreg.basic.value == "del" && itemreg.imgbasic.value =="") {
        alert("�⺻�̹����� �ʼ��Դϴ�.");
        return;
    } else {
        if (itemreg.imgbasic.value != "") {
            if (CheckImage(itemreg.imgbasic, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
                return;
            }
        }
    }

    if (itemreg.imgmask.value != "") {
        if (CheckImage(itemreg.imgmask, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd1.value != "") {
        if (CheckImage(itemreg.imgadd1, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd2.value != "") {
        if (CheckImage(itemreg.imgadd2, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd3.value != "") {
        if (CheckImage(itemreg.imgadd3, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd4.value != "") {
        if (CheckImage(itemreg.imgadd4, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if (itemreg.imgadd5.value != "") {
        if (CheckImage(itemreg.imgadd5, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40) != true) {
            return;
        }
    }

    if(confirm("�����Ͻðڽ��ϱ�?") == true){
        itemreg.submit();
    }

}

function CloseWindow() {
    // opener.focus();
    window.close();
}


// ============================================================================
// �̹���ǥ��
function ClearImage(img,fsize,wd,ht) {
 // 	img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this.form." + img.name + ", <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg', "+ fsize +");\" class='text' size='"+ fsize +"'>";
 
    document.getElementById("div"+ img.name).style.display = "none";

	var e = eval("itemreg."+img.name.substr(3,img.name.length));
	e.value = "del";
}

function ClearImage2(img,fsize,wd,ht,num) {
	var imgcnt = $('input[name="addimgname"]').length;
    img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +", "+num+");\" class='text' size='"+ fsize +"'>";
	$("#divaddimgname"+(num+1)+"").hide();
	
	if(imgcnt > 1){
    	document.itemreg.addimgdel[num].value = "del";
    }else{
    	document.itemreg.addimgdel.value = "del";
    }
}

function ClearImage3(img,fsize,wd,ht,num) {
	var imgcnt = $('input[name="addmobileimgname"]').length;
    img.outerHTML = "<input type='file' name='" + img.name + "' onchange=\"CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +", "+num+");\" class='text' size='"+ fsize +"'>";
	$("#divaddmobileimgname"+(num+1)+"").hide();
	
	if(imgcnt > 1){
    	document.itemreg.addmobileimgdel[num].value = "del";
    }else{
    	document.itemreg.addmobileimgdel.value = "del";
    }
}

function oldClearImage(img,fsize,wd,ht) {
	$("#divimg"+img+"").hide();
	$("input[name='"+img+"']").val("del");
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
        alert("�̹��������� ������ ���ϸ� ����ϼ���.[" + extname + "]");
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
        alert("�̹��������� ������ ���ϸ� ����ϼ���.[" + extname + "]");
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


//��ǰ�����̹����߰�
function InsertImageUp() {
	var f = document.all;
	var rowLen = f.imgIn.rows.length;

	if(rowLen > 6){
		alert("�̹����� �ִ� 7������ �����մϴ�.");
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
	c0.innerHTML = '��ǰ�����̹��� #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 800, 1600, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40, '+parseInt(rowLen-1)+');" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' �̹��������" class="button" onClick="ClearImage2(this.form.addimgname['+parseInt(rowLen-1)+'],40, 800, 1600, '+parseInt(rowLen-1)+')"> (����,800X1600, Max 800KB,jpg,gif)';
	c1.innerHTML += '<input type="hidden" name="addimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addimgdel" value="">';
}

//����ϻ�ǰ���̹����߰�
function InsertMobileImageUp() {
	var f = document.all;
	var rowLen = f.MobileimgIn.rows.length;

	if(rowLen > 11){
		alert("�̹����� �ִ� 12������ �����մϴ�.");
		return;
	}
	
	var i = rowLen;
	var r  = f.MobileimgIn.insertRow(rowLen++);
	var c0 = r.insertCell(0);
	var c1 = r.insertCell(1);

	r.style.textAlign = 'left';
	c0.style.height = '30';
	c0.style.width = '15%';
	c0.style.background = '#DDDDFF';
	c0.innerHTML = '����ϻ�ǰ���̹��� #' + rowLen + ' :';
	c1.style.background = '#FFFFFF';
	c1.innerHTML = '<input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, '+String.fromCharCode(39)+'jpg,gif'+String.fromCharCode(39)+',40);" class="text" size="40"> ';
	c1.innerHTML += '<input type="button" value="#'+parseInt(rowLen)+' ������̹��������" class="button" onClick="ClearImage3(this.form.addmobileimgname['+parseInt(rowLen-1)+'],40, 640, 1200, '+parseInt(rowLen-1)+')"> (����,640X1200, Max 400KB,jpg,gif)';
	c1.innerHTML += '<input type="hidden" name="addmobileimggubun" value="'+parseInt(rowLen)+'">';
	c1.innerHTML += '<input type="hidden" name="addmobileimgdel" value="">';
}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left">
		<br>�̹�������
		<br>- �ٹ����ٿ��� �̹����� ����� ��� ���� �Է����� ���ñ� �ٶ��ϴ�.
		<br>- �̹����� <font color=red><%= CBASIC_IMG_MAXSIZE %>kb</font> ���� �ø��� �� �ֽ��ϴ�.
		<br>&nbsp;&nbsp;(�̹�������� <font color=red>���μ������� ������</font>�� �԰ݿ� ���� �ʰ� ������ּ���. �԰��ʰ��� ����� ���� �ʽ��ϴ�.)
		<br>- <font color=red>����޿��� Save For Web����, Optimizeüũ, ������ 80%����</font>�� ����� �� �÷��ֽñ� �ٶ��ϴ�.
		<br><br>�̹��� ������ <font color=red>CTRL + F5 (��Ʈ�� F5 ��ư)</font></a> �����ž� ������ �̹����� Ȯ���Ͻ� �� �ֽ��ϴ�
		<br><br><input type="button" value=" ���ΰ�ħ " class="button" onClick="location.reload();"> &nbsp;&nbsp; <input type="button" value=" �� �� " class="button" onClick="CloseWindow()"><br>&nbsp;
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/itemImageModify.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�⺻�̹��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fbasicimage <> "") then %>
		<div id="divimgbasic" style="display:block;">
		<img src="<%= oitem.FOneItem.Fbasicimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgbasic" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgbasic" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgbasic,40, 1000, 1000)"> (<font color=red>�ʼ�</font>,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="basic">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�������̹���(�ڵ�����) :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<% if (oitem.FOneItem.Ficon1image <> "") then %>
		<img src="<%= oitem.FOneItem.Ficon1image %>" width="200" height="200">
	<% end if %>
	<% if (oitem.FOneItem.Ficon2image <> "") then %>
		<img src="<%= oitem.FOneItem.Ficon2image %>" >
	<% end if %>
	<% if (oitem.FOneItem.Flistimage120 <> "") then %>
		<img src="<%= oitem.FOneItem.Flistimage120 %>" width="120" height="120">
	<% end if %>
	<% if (oitem.FOneItem.Flistimage <> "") then %>
		<img src="<%= oitem.FOneItem.Flistimage %>" width="100" height="100">
	<% end if %>
	<% if (oitem.FOneItem.Fsmallimage <> "") then %>
		<img src="<%= oitem.FOneItem.Fsmallimage %>" width="50" height="50">
	<% end if %>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">����(����)�̹��� :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmaskimage <> "") then %>
		<div id="divimgmask" style="display:block;">
		<img src="<%= oitem.FOneItem.Fmaskimage %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgmask" style="display:none;"></div>
	  <% end if %>
	  <input type="file" name="imgmask" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg', 40);" class="text" size="40">
	  <input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgmask,40, 1000, 1000)"> (����,1000X1000,<b><font color="red">jpg</font></b>)
	  <input type="hidden" name="mask">
	</td>
</tr>
<tr height="1" bgcolor="#CCCCCC"><td colspan="4"></td></tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,1) <> "") then %>
		<div id="divimgadd1" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,1) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd1" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd1" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd1,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add1">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���2 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,2) <> "") then %>
		<div id="divimgadd2" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,2) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd2" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd2" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd2,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���3 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,3) <> "") then %>
		<div id="divimgadd3" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,3) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd3" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd3" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd3,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add3">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���4 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,4) <> "") then %>
		<div id="divimgadd4" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,4) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd4" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd4" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd4,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add4">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">�߰��̹���5 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitemAddImage.GetImageAddByIdx(0,5) <> "") then %>
		<div id="divimgadd5" style="display:block;">
		<img src="<%=oitemAddImage.GetImageAddByIdx(0,5) %>" width="300" height="300">
		</div>
	  <% else %>
	  <div id="divimgadd5" style="display:none;"></div>
	  <% end if %>
		<input type="file" name="imgadd5" onchange="CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40);" class="text" size="40">
		<input type="button" value="�̹��������" class="button" onClick="ClearImage(this.form.imgadd5,40, 1000, 1000)"> (����,1000X1000,jpg,gif)
		<input type="hidden" name="add5">
	</td>
</tr>
</table>
<%
	Dim cImg, k, vArr, j
	set cImg = new CItemAddImage
	cImg.FRectItemID = itemid
	vArr = cImg.GetAddImageListIMGTYPE1
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="imgIn">
	<% If isArray(vArr) Then
			If vArr(3,UBound(vArr,2)) > 0 Then
			For k = 1 To vArr(3,UBound(vArr,2))
	%>
			  <tr align="left">
			  	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #<%= (k) %> :</td>
			  	<td bgcolor="#FFFFFF">
		  		<%
		  		If cImg.IsImgExist(vArr,k) Then
		    		For j = 0 To UBound(vArr,2)
		    			If CStr(vArr(3,j)) = CStr(k) AND (vArr(4,j) <> "" and isNull(vArr(4,j)) = False) Then
							Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vArr(1,j)) & "/" & vArr(4,j) & """ height=""250""></div>"
							Exit For
		    			End If
		    		Next
				Else
					Response.Write "<div id=""divaddimgname"&(k)&""" style=""display:none;""></div>"
				End If
				%>
			      <input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40, <%= (k-1) %>);" class="text" size="40">
			      <input type="button" value="#<%= (k) %> �̹��������" class="button" onClick="ClearImage2(this.form.addimgname<%=CHKIIF(vArr(3,UBound(vArr,2))=1,"","["&(k-1)&"]")%>,40, 1000, 1000, <%= (k-1) %>)"> (����,800X1600, Max 800KB,jpg,gif)
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
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #1 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname1" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40,0);" class="text" size="40">
				<input type="button" value="#1 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[0],40, 1000, 1000, 0)"> (����,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="1">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #2 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname2" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40,1);" class="text" size="40">
				<input type="button" value="#2 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[1],40, 1000, 1000, 1)"> (����,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="2">
				<input type="hidden" name="addimgdel" value="">
			</td>
		</tr>
		<tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #3 :</td>
			<td bgcolor="#FFFFFF">
				<div id="divaddimgname3" style="display:none;"></div>
				<input type="file" name="addimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 1000, 1000, 'jpg,gif',40,2);" class="text" size="40">
				<input type="button" value="#3 �̹��������" class="button" onClick="ClearImage2(this.form.addimgname[2],40, 1000, 1000, 2)"> (����,800X1600, Max 800KB,jpg,gif)
				<input type="hidden" name="addimggubun" value="3">
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
      <input type="button" value="��ǰ�����̹����߰�" class="button" onClick="InsertImageUp()">
      <font color="red">* ���ε尡 �� �̹����� ����� �ȳ����� ���ΰ�ħ(CTRL + F5(��Ʈ�� F5 ��ư))�� ���ּ���.</font>
  	</td>
  </tr>
</table>


<%
	Dim cmImg, mk, vmArr, mj
	set cmImg = new CItemAddImage
	cmImg.FRectItemID = itemid
	vmArr = cmImg.GetAddImageListIMGTYPE2
%>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" id="MobileimgIn">
	<% If isArray(vmArr) Then
			If vmArr(3,UBound(vmArr,2)) > 0 Then
			For mk = 1 To vmArr(3,UBound(vmArr,2))
	%>

			  <tr align="left">
				<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #<%= (mk) %> :</td>
				<td bgcolor="#FFFFFF">
				<%
				If cmImg.IsImgExist(vmArr,mk) Then
					For mj = 0 To UBound(vmArr,2)
						If CStr(vmArr(3,mj)) = CStr(mk) AND (vmArr(4,mj) <> "" and isNull(vmArr(4,mj)) = False) Then
							Response.Write "<div id=""divaddmobileimgname"&(mk)&""" style=""display:block;""><img src=""" & webImgUrl & "/item/contentsimage/" & GetImageSubFolderByItemid(vmArr(1,mj)) & "/" & vmArr(4,mj) & """ height=""250""></div>"
							Exit For
						End If
					Next
				Else
					Response.Write "<div id=""divaddmobileimgname"&(mk)&""" style=""display:none;""></div>"
				End If
				%>
				  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40, <%= (mk-1) %>);" class="text" size="40">
				  <input type="button" value="#<%= (mk) %> �̹��������" class="button" onClick="ClearImage3(this.form.addmobileimgname<%=CHKIIF(vmArr(3,UBound(vmArr,2))=1,"","["&(mk-1)&"]")%>,40, 640, 1200, <%= (mk-1) %>)"> (����,640X1200, Max 400KB,jpg,gif)
				  <input type="hidden" name="addmobileimggubun" value="<%= (mk) %>">
				  <input type="hidden" name="addmobileimgdel" value="">
				</td>
			  </tr>
	<%
			Next
			End IF
		Else
	%>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #1 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#1 �̹��������" class="button" onClick="ClearImage3(this.form.addmobileimgname[0],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="1">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #2 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#2 �̹��������" class="button" onClick="ClearImage3(this.form.addmobileimgname[1],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="2">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
		  <tr align="left">
			<td height="30" width="15%" bgcolor="#DDDDFF">����ϻ�ǰ���̹��� #3 :</td>
			<td bgcolor="#FFFFFF">
			  <input type="file" name="addmobileimgname" onchange="CheckImage2(this, <%= CBASIC_IMG_MAXSIZE %>, 640, 1200, 'jpg,gif',40);" class="text" size="40">
			  <input type="button" value="#3 �̹��������" class="button" onClick="ClearImage3(this.form.addmobileimgname[2],40, 640, 1200)"> (����,640X1200, Max 400KB,jpg,gif)
				<input type="hidden" name="addmobileimggubun" value="3">
				<input type="hidden" name="addmobileimgdel" value="">
			</td>
		  </tr>
	<%
	   End IF %>
</table>
<%	set cmImg = nothing %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
 	<td colspan="4">
 	<font color="red"><strong>�� ����� ��ǰ�� �̹����� ������ �� �������� ��ü �˴ϴ�. html�� ������� ���� �����̿��� �������� ���ε� ���ֽñ� �ٶ��ϴ�.<br>�� ����� ��ǰ�󼼿��� �̹����� �߶� �÷��ֽñ� �ٶ��ϴ�.</strong></font>
 	</td>
 </tr>
  <tr align="left">
  	<td bgcolor="#FFFFFF">
      <input type="button" value="����ϻ�ǰ���̹����߰�" class="button" onClick="InsertMobileImageUp()">
      <font color="red">* ���ε尡 �� �̹����� ����� �ȳ����� ���ΰ�ħ(CTRL + F5(��Ʈ�� F5 ��ư))�� ���ּ���.</font>
  	</td>
  </tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #1 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage <> "") then %>
		<div id="divimgmain" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="�̹��������" class="button" onClick="oldClearImage('main', 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #2:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage2 <> "") then %>
		<div id="divimgmain2" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage2 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain2" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="�̹��������" class="button" onClick="oldClearImage('main2', 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main2">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">��ǰ�����̹��� #3:</td>
	<td bgcolor="#FFFFFF" colspan="3">
	  <% if (oitem.FOneItem.Fmainimage3 <> "") then %>
		<div id="divimgmain3" style="display:block;">
		<img src="<%=oitem.FOneItem.Fmainimage3 %>" width="400">
		</div>
	  <% else %>
	  <div id="divimgmain3" style="display:none;"></div>
	  <% end if %>
		<input type="button" value="�̹��������" class="button" onClick="oldClearImage('main3', 40, 800, 1600)"> (����,800X1600, Max <%= CMAIN_IMG_MAXSIZE %>KB,jpg,gif)
		<input type="hidden" name="main3">
	</td>
</tr>
</table>
</form>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="�����ϱ�" class="button" onClick="SubmitSave()">
          <input type="button" value="����ϱ�" class="button" onClick="CloseWindow()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<%
set oitem = Nothing
set oitemAddImage = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->