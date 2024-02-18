<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->
<%

dim yearUse : yearUse = request("yearUse")
dim DiaryType : DiaryType = request("DiaryType")

dim page : page=1
dim pagesize : pagesize=20
dim mdiary,i


dim arrPick
set mdiary = new ClsDiary
mdiary.FYearUse =YearUse
mdiary.FDiaryType=DiaryType
mdiary.FCurrPage= page
mdiary.FPageSize=pagesize
mdiary.FScrollCount=10
arrPick = mdiary.getMdsPickList

%>
<script language="javascript">
function fnDelListitem(num){
	var tbl = document.getElementById('regtbl');
	tbl.deleteRow(num);

	fnReSortList();
}
function fnReSortList(){

	var tar = document.getElementsByName('rank');

	for(var i=0;i<tar.length;i++){
		tar[i].value=i+1;
	}
}
function showimage(img){
	var pop = window.open('viewImage.asp?imageUrl='+img,'imgview','width=600,height=600,resizable=yes');
}
window.resizeTo(800,650);

function subfrm(){
	document.regfrm.submit();
}
function resetfrm(){
	document.regfrm.reset();
}
</script>
<!-- 상단 메뉴 -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" valign="top">
		<iframe src="inc_diary_mdspick_itemlist.asp?yearuse=<%= yearUse %>" name="itemlistframe" frameborder="0" width="100%" height="700"></iframe>
	</td>
	<td width="10" valign="top" align="center" style="padding-top:70">

	</td>
	<td valign="top" width="350">
		<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			<tr height="10" valign="bottom">
		        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		        <td background="/images/tbl_blue_round_02.gif"></td>
		        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
			</tr>
			<tr valign="top" style="padding : 0 0 10 0">
		        <td background="/images/tbl_blue_round_04.gif"></td>
		        <td align="right">
		        	&nbsp;</td>
				<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
			</tr>
		</table>

		<table width="350" border="0" cellpadding="0" cellspacing="1"  class="a" id="regtbl" bgcolor="<%= adminColor("tablebg") %>">
			<form name="regfrm" method="post" action="Proc_diary_mdsPick.asp">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td width="50" align="center">순위 </td>
				<td width="60" align="center">번호</td>
				<td width="70" align="center">상품번호</td>
				<td width="50" align="center">이미지</td>
				<td width="50" align="center">삭제</td>
				<td></td>
			</tr>
			<% if isArray(arrPick) then %>
			<% for i=0 to Ubound(arrPick,2) %>
			<tr bgcolor="#FFFFFF">
				<td align="center"><input type="text" name="rank" size="3" value="<%= arrPick(1,i) %>"> </td>
				<td align="center"><input type="text" name="diaryid" size="5" value="<%= arrPick(0,i) %>"> </td>
				<td align="center"><input type="text" name="itemid" size="7" value="<%= arrPick(2,i) %>"> </td>
				<td align="center"><img src="<%= "http://webimage.10x10.co.kr/diary_collection/" & yearUse & "/icon/" & arrPick(3,i) %>" width="25" height="25"> </td>
				<td align="center"><span onclick="fnDelListitem(parentElement.parentElement.rowIndex);" style="cursor:pointer">[X]</span></td>
				<td></td>
			</tr>
			<% next %>
			<% end if %>

			</form>
		</table>

		<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
		    <tr valign="bottom" height="25">
		        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
		        <td valign="bottom" align="center">
		        	<input type="button" class="button" value="적용" onclick="subfrm();">&nbsp;&nbsp;&nbsp;
		        	<input type="button" class="button" value="취소" onclick="window.close();">
		        </td>
		        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
		    </tr>
		    <tr valign="top" height="10">
		        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		        <td background="/images/tbl_blue_round_08.gif"></td>
		        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
		    </tr>
		</table>
	<!--<iframe name="pickframe" frameborder="1" width="300" height="100%"></iframe>-->
	</td>
</tr>
</table>
<form name="pagingFrm" method="get" action="?">
<input type="hidden" name="page" value="" />
<input type="hidden" name="yearuse" value="<%= YearUse %>">
<input type="hidden" name="DiaryType" value="<%= DiaryType %>" />
<input type="hidden" name="menupos" value="<%= menupos %>" />
</form>
<% set mdiary = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->