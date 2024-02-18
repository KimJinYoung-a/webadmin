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


dim objInfo ,intLoop
set objInfo = new clsDiary
objinfo.FYearUse = YearUse
objInfo.getDiaryInfo idx


%>
<script language="javascript" type="text/javascript">

function delItem(info_idx,info_gubun){
	document.delFrm.info_idx.value=info_idx;
	document.delFrm.info_gubun.value=info_gubun;
	document.delFrm.submit();
}

function checkPageCnt(str) {

	if(str.value < '0' || str.value.length < 0){
		str.value='0';
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
	document.imginputfrm.action='diary_img_input.asp';
	document.imginputfrm.submit();
}

function jsImgDel(divnm,iptNm,vPath){

	window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.action='http://upload.10x10.co.kr/linkweb/diary_collection/diary_collection_image_del_proc.asp';
	document.imginputfrm.submit();
}

function subchk(){

	var infoname ='';
	var infogubun ='';
	var infoImage ='';
	var infocnt ='';
	for(var i=0;i<document.regfrm.elements.length;i++){


		if(document.regfrm.elements[i].name.substr(0,9)=="info_name"){
			infoname = infoname + document.regfrm.elements[i].value + ',';
		}
		if(document.regfrm.elements[i].name.substr(0,9)=="infogubun"){
			infogubun = infogubun + document.regfrm.elements[i].value + ',';
		}

		if(document.regfrm.elements[i].name.substr(0,7)=="infoimg"){
			infoImage = infoImage + document.regfrm.elements[i].value + ',';
		}
		if(document.regfrm.elements[i].name.substr(0,12)=="info_pageCnt"){
			infocnt = infocnt + document.regfrm.elements[i].value + ',';
		}
	}
	document.realregfrm.mode.value=document.regfrm.mode.value;
	document.realregfrm.infoname.value=infoname;
	document.realregfrm.infogubun.value= infogubun;
	document.realregfrm.infoImage.value=infoImage;
	document.realregfrm.infocnt.value=infocnt;

	document.realregfrm.TotalPageName.value	= document.regfrm.TotalPageName.value;
	document.realregfrm.TotalPagepageCnt.value	= document.regfrm.TotalPagepageCnt.value;
	document.realregfrm.etcname.value		= document.regfrm.etcname.value;


	document.realregfrm.submit();
}

document.domain="10x10.co.kr"

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
        	<b>내지 구성 등록 </b></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 등록 부분 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="">
	<input type="hidden" name="mode" value="edit">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="100">내지구성</td>
		<td align="center" >이미지</td>
		<td></td>
		<td align="center" width="100">Pages</td>
	</tr>
<% if objInfo.FResultCount>0 then %>
	<% for intLoop =0 to objInfo.FResultCount -1 %>
	<%if objInfo.FItemList(intLoop).FinfoGubun>="1" and objInfo.FItemList(intLoop).FinfoGubun<=14 then %>

	<tr bgcolor="#FFFFFF">
		<td><input type="text" name="info_name" value="<%= objInfo.FItemList(intLoop).Finfoname %>"></td>
		<td>
			<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('infodiv<%= intLoop %>','infoimg<%= intLoop %>','info','200','400','false');"/>
			<input type="button" class="button" size="30" value="이미지 삭제" onclick="jsImgDel('infodiv<%= intLoop %>','infoimg<%= intLoop %>','info');"/>

			<input type="hidden" name="infogubun" value="<%= objInfo.FItemList(intLoop).FinfoGubun %>">
			<input type="hidden" name="infoimg<%= intLoop %>" value="<%= objInfo.FItemList(intLoop).Finfoimg %>">
		</td>
		<td>
			<div align="center" id="infodiv<%= intLoop %>">
				<% if (not isnull(objInfo.FItemList(intLoop).Finfoimg)) and (trim(objInfo.FItemList(intLoop).Finfoimg)<>"") then%>
				<img src="<%= objInfo.FItemList(intLoop).getInfoImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= objInfo.FItemList(intLoop).getInfoImgUrl %>');">
				<% end if %>
			</div>
		</td>
		<td><input type="text" name="info_pageCnt" size="3" value="<%= objInfo.FItemList(intLoop).FinfoPageCnt %>"  />장</td>
	</tr>

	<% end if %>
	<%if objInfo.FItemList(intLoop).FinfoGubun="15" then %>
	<!-- TotalPages -->
	<tr bgcolor="#FFFFFF">
		<td><input type="text" name="TotalPageName" value="<%= objInfo.FItemList(intLoop).Finfoname %>"></td>
		<td colspan="2">&nbsp;</td>
		<td><input type="text" name="TotalPagepageCnt" size="3" value="<%= objInfo.FItemList(intLoop).FinfoPageCnt %>" />장</td>
	</tr>
	<% end if %>

	<%if objInfo.FItemList(intLoop).FinfoGubun="16" then %>
	<!-- 기타 내용 입력 -->
	<tr bgcolor="#FFFFFF">
		<td align="center">기타 내용</td>
		<td colspan="3"><textarea name="etcname" cols="50" rows="5"><%= objInfo.FItemList(intLoop).Finfoname %></textarea></td>
	</tr>
	<% end if %>

	<% next %>
<% end if %>
	</form>
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
<form name="realregfrm" method="post" action="diary_reg_Info_proc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="infoname" value="">
<input type="hidden" name="infogubun" value="">
<input type="hidden" name="infoImage" value="">
<input type="hidden" name="infocnt" value="">

<input type="hidden" name="TotalPageName" value="">
<input type="hidden" name="TotalPagepageCnt" value="">
<input type="hidden" name="etcname" value="">
</form>
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
<% set objInfo = nothing %>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->
