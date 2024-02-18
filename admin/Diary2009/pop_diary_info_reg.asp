<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%
dim mode,diaryid
mode=request("mode")
diaryid= request("diaryid")

dim objInfo ,intLoop
set objInfo = new DiaryCls
objInfo.getDiaryInfo diaryid


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
	var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
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
	document.imginputfrm.action='PopImgInput.asp';
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

function subchk(diaryid,idx,image_count , contents_idx){
	var infoImage ='';
	var infocnt ='';

	for(var i=0;i<document.regfrm.elements.length;i++){
		if(document.regfrm.elements[i].name.substr(0,7)=="infoimg"){
			infoImage = infoImage +document.regfrm.elements[i].value+',' ;
		}
		if(document.regfrm.elements[i].name.substr(0,12)=="info_pageCnt"){
			infocnt = infocnt + document.regfrm.elements[i].value + ',';
		}

	}
	document.realregfrm.mode.value=document.regfrm.mode.value;
	document.realregfrm.infoImage.value=infoImage;
	document.realregfrm.idx.value= idx;
	document.realregfrm.contents_idx.value= contents_idx;
	document.realregfrm.image_count.value =	image_count;
	document.realregfrm.infocnt.value = infocnt;
	document.realregfrm.submit();
}

document.domain="10x10.co.kr"

</script>

<!-- 등록 부분 -->
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#9d9d9d">
	<form name="regfrm" method="post" action="">
	<input type="hidden" name="mode" value="edit">
	<tr bgcolor="#FFFFFF">
		<td align="center" width="100">내지구성</td>
		<td align="center" >이미지</td>
		<td></td>
		<td align="center" width="100">Pages</td>
		<td align="center">비고</td>
	</tr>
<% if objInfo.FResultCount>0 then %>
	<% for intLoop =0 to objInfo.FResultCount -1 %>

	<tr bgcolor="#FFFFFF">
		<td><input type="text" name="info_name" value="<%= objInfo.FItemList(intLoop).foption_value %>"></td>
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
		<td><input type="button" class="button" value="저장" onclick="subchk('<%= diaryid %>','<%= objInfo.FItemList(intLoop).fidx %>','<%=intLoop%>','<%= objInfo.FItemList(intLoop).Finfoname %>');"/></td>
	</tr>

	<% next %>
<% end if %>
	</form>

<form name="realregfrm" method="post" action="proc_diary_Info.asp">
<input type="hidden" name="diaryid" value="<%= diaryid %>">
<input type="hidden" name="infoImage" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="infocnt" value="">
<input type="hidden" name="image_count" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="infogubun" value="">
<input type="hidden" name="contents_idx">
<input type="hidden" name="TotalPageName" value="">
<input type="hidden" name="TotalPagepageCnt" value="">
<input type="hidden" name="etcname" value="">
</form>
<form name="imginputfrm" method="post" action="">
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