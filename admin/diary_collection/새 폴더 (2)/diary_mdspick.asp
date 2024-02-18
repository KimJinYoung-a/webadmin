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

set mdiary = new ClsDiary
mdiary.FYearUse =YearUse
mdiary.FDiaryType=DiaryType
mdiary.FCurrPage= page
mdiary.FPageSize=pagesize
mdiary.FScrollCount=10
mdiary.GetDiaryList

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
	<td width="350" valign="top"><iframe src="diary_mdspick_itemList.asp?yearuse=<%= yearUse %>" name="itemlistframe" frameborder="0" width="350" height="700"></iframe></td>
	<td width="70" valign="top" align="center" style="padding-top:70">
		<input type="button" class="button" value="==>"><br><br>
		<input type="button" class="button" value="<==">
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
			<form name="regfrm" method="post" action="diary_mdsPick_proc.asp">
			<tr bgcolor="<%= adminColor("tabletop") %>">
				<td width="50" align="center">순위 </td>
				<td width="60" align="center">번호</td>
				<td width="70" align="center">상품번호</td>
				<td width="50" align="center">이미지</td>
				<td width="50" align="center">삭제</td>
				<td></td>
			</tr>

			</form>
		</table>

		<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
		    <tr valign="bottom" height="25">
		        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
		        <td valign="bottom" align="center">
		        	<input type="button" class="button" value="적용" onclick="subfrm();">&nbsp;&nbsp;&nbsp;
		        	<input type="button" class="button" value="취소" onclick="resetfrm();">
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
