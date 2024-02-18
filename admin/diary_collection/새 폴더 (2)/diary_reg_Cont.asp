<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->



<%
dim mode,idx

mode=request("mode")
idx= request("idx")

dim objDiary ,YearUse
set objDiary = new clsDiary
objDiary.getDiaryItem idx
YearUse = objDiary.DiaryPrd.FYear
set objDiary = nothing

dim objCon,intLoop
set objCon = new clsDiary
objCon.FYearUse = YearUse
objCon.getDiaryContens idx



%>
<script type="text/javascript" language="javascript">

function editCont(id){
	document.regfrm.cont_idx.value=id;
	document.regfrm.mode.value='modify';
	document.regfrm.cont_text.value=document.getElementById('txt_'+id).value;
	document.regfrm.submit();
}

function delCont(id){
	var msg = confirm('삭제를 하시면 이미지와 내용이 같이 삭제가 됩니다.\n삭제 하시겠습니까?');
	if (msg) {
		document.regfrm.cont_idx.value=id;
		document.regfrm.mode.value='del';
		document.regfrm.submit();
	}
}

function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes');
}
function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

	window.open('','imginput','width=10,height=10,menubar=no,toolbar=no,scrollbars=no,status=no,resizable=no,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='diary_img_input.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=10,height=10,menubar=no,toolbar=no,scrollbars=no,status=no,resizable=no,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imgdel';
	document.imginputfrm.action='http://upload.10x10.co.kr/linkweb/diary_collection/diary_collection_image_del_proc.asp';
	document.imginputfrm.submit();
}

function subchk(){
	document.regfrm.submit();
}

document.domain="10x10.co.kr"
window.resizeTo(600,800);
</script>
<!-- 상단 메뉴 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="center">
        	<b>상세 설명 이미지 등록 </b><br>내용에 관한 수정만 가능합니다.<br>이미지를 수정하실땐 삭제후 다시 등록하셔야 합니다</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<!-- 등록된 이미지 리스트 -->

<% If objCon.FResultCount >0 Then%>
	<%For intLoop=0 to objCon.FResultCount -1 %>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center">이미지</td>
		<td>
		<% if objCon.FItemList(intLoop).ConfImg<>"" then %>
			<img src="<%= objCon.FItemList(intLoop).getContImgUrl %>" width="100" height="100" onclick="showimage('<%= objCon.FItemList(intLoop).getContImgUrl %>');" style="cursor:pointer" />&nbsp;&nbsp;
		<% end if %>
			<input type="button" class="button" value="수정" onclick="editCont('<%= objCon.FItemList(intLoop).ConIdx %>');">&nbsp;&nbsp;
			<input type="button" class="button" value="삭제" onclick="delCont('<%= objCon.FItemList(intLoop).ConIdx %>');">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center">내용</td>
		<td><textarea id="txt_<%= objCon.FItemList(intLoop).ConIdx %>" cols="60" rows="10"><%= objCon.FItemList(intLoop).ConTTxt %></textarea></td>
	</tr>
	<tr>
		<td colspan="2" height="5" bgcolor="#CCCCCC"></td>
	</tr>
	<% next %>
<% Else %>
	<tr bgcolor="#FFFFFF">
		<td  align="center" height="100"> [등록된 이미지가 없습니다.이미지를 등록하여 주세요.] </td>
	</tr>
<% End If %>

</table>
<!-- 이미지 등록 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="diary_reg_Cont_proc.asp">
	<input type="hidden" name="cont_idx" value="">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="mode" value="write">
	<tr>
		<td colspan="2" align="center"  bgcolor="<%= adminColor("tablebg") %>">상세 설명 추가</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">상세 이미지</td>
		<td>
			<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('contimg','contImgName','cont','200','600','false');"/>
			<input type="button" class="button" size="30" value="이미지 삭제" onclick="jsImgDel('contimg','contImgName','cont');"/>
			<input type="hidden" name="contImgName" value="">
			<div id="contimg"></div></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">상세 텍스트</td>
		<td><textarea name="cont_text" cols="55" rows="10" ></textarea></td>
	</tr>
	</form>
</table>
<!-- 하단  시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<input type="button" class="button" value="확인" onclick="subchk();"/>&nbsp;&nbsp;
			<input type="button" class="button" value="취소" />
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<form name="imginputfrm" method="post" action="">
<input type="hidden" name="YearUse" value="<%= YearUse %>">
<input type="hidden" name="divName" value="">
<input type="hidden" name="orgImgName" value="">
<input type="hidden" name="inputname" value="">
<input type="hidden" name="ImagePath" value="">
<input type="hidden" name="maxFileSize" value="">
<input type="hidden" name="maxFileWidth" value="">
<input type="hidden" name="makeThumbYn" value="">
</form>
</body>
</html>
<% set objCon= nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
