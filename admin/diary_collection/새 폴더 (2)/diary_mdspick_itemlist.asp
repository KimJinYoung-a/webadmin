<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<%

dim yearUse : yearUse = request("yearUse")
dim DiaryType : DiaryType = request("DiaryType")

dim page : page=1
dim pagesize : pagesize=20
dim mdiary,i

set mdiary = new ClsDiary
mdiary.FYearUse =YearUse
mdiary.FDiaryType=DiaryType
mdiary.FCurrPage= page
mdiary.FPageSize=pagesize
mdiary.FScrollCount=10
mdiary.GetDiaryList

%>
<script language="javascript">
//분류별 검색
function FnSelDiaryType(varDiaryType){
	document.pagingFrm.page.value=1;
	document.pagingFrm.DiaryType.value=varDiaryType;
	document.pagingFrm.submit();
}
function fnpicktmp(rk,di,it,img){

	var tmp =1;
	var rank = parent.document.getElementsByName("rank");
	var diaryid = parent.document.getElementsByName("diaryid");

	if(rank.length>9){
		alert('10개의 상품까지만 등록 가능합니다.');
		return false;
	}
	if(rank.length<1){
		tmp =1;
	} else {
		for (i=0;i<rank.length;i++){

			tmp = (tmp > rank[i].value ? tmp:rank[i].value);

			if (eval(di)==eval(diaryid[i].value)){
				alert('중복된 상품 입니다.');
				return false;
			s}
		}
		tmp= eval(tmp) +1;
	}




	var tbl = parent.document.getElementById("regtbl");
	var oRow = tbl.insertRow();
	var oCell1 = oRow.insertCell();
	var oCell2 = oRow.insertCell();
	var oCell3 = oRow.insertCell();
	var oCell4 = oRow.insertCell();
	var oCell5 = oRow.insertCell();
	oRow.bgColor='#FFFFFF';

	oCell1.align="center";
	oCell2.align="center";
	oCell3.align="center";
	oCell4.align="center";
	oCell5.align="center";

	oCell1.innerHTML ='<input type="text" name="rank" size="3" value="' + tmp + '">';
	oCell2.innerHTML ='<input type="text" name="diaryid" size="5" value="' + di + '">';
	oCell3.innerHTML ='<input type="text" name="itemid" size="7" value="' + it + '">';
	oCell4.innerHTML ='<img src="' + img + '" width="25" height="25" onclick="showimage(\'' +img+'\')" style="cursor:pointer">';
	oCell5.innerHTML ='<span onclick="fnDelListitem(parentElement.parentElement.rowIndex);" style="cursor:pointer">[X]</span>';

	//parent.document.regfrm.rank.value=rk;
	//parent.document.regfrm.diaryid.value=di;
	//parent.document.regfrm.itemid.value=it;

}
</script>
<!-- 상단 메뉴 -->

<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="right">
        	<select name="DiaryType"  onchange="FnSelDiaryType(this.value);">
				<option value="" 		 <% if DiaryType="" 		then response.write "selected"  %>>전체</option>
				<option value="illust" <% if DiaryType="illust" then response.write "selected"  %>>일러스트</option>
				<option value="photo"  <% if DiaryType="photo"  then response.write "selected"  %>>포토/명화</option>
				<option value="system" <% if DiaryType="system" then response.write "selected"  %>>시스템</option>
			</select></td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 중간 메인부분 -->
<table width="350" border="0" cellpadding="0" cellspacing="1"  class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center">번호</td>
		<td width="60" align="center">구분</td>
		<td width="60" align="center">이미지</td>
		<td width="50" align="center">상품번호</td>
		<td width="250" align="center">상품명</td>

	</tr>
	<!-- 리스트 표시 -->
	<% if mdiary.FResultCount<=0 then %>
		<!-- 없음 -->
	<% else %>
	<% for i = 0 to mdiary.FResultCount-1 %>

		<% if mdiary.FItemList(i).FIsusing="N" then %>
		<tr bgcolor="#ECECEC">
		<% else %>
		<tr bgcolor="#FFFFFF">
		<% end if %>

		<td align="center">
			<span onclick="fnpicktmp('1','<%= mdiary.FItemList(i).FIdx %>','<%= mdiary.FItemList(i).Fitemid %>','<%= db2html(mdiary.FItemList(i).getListImgUrl) %>');" style="cursor:pointer">
				<%= mdiary.FItemList(i).FIdx %></span>
		</td>
		<td align="center"><%= mdiary.FItemList(i).StrDiaryTypeName %></td>
		<td align="center"><img src="<%= db2html(mdiary.FItemList(i).getListImgUrl) %>"  border="0"></td>
		<td align="center"><%= mdiary.FItemList(i).Fitemid %></td>
		<td align="center"><%= db2html(mdiary.FItemList(i).FItemName) %></td>

	</tr>
<% next %>
</table>
<% end if %>

<!-- 하단 페이징 시작 -->
<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif">&nbsp;</td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<form name="pagingFrm" method="get" action="?">
<input type="hidden" name="page" value="" />
<input type="hidden" name="yearuse" value="<%= YearUse %>">
<input type="hidden" name="DiaryType" value="<%= DiaryType %>" />

</form>
<% set mdiary = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
