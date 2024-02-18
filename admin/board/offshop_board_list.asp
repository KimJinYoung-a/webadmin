<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshop_boardcls.asp" -->
<%

dim i, j, page

page = request("page")
if page="" then page=1

'==============================================================================
'나의 1:1질문답변
dim offnews
set offnews = New COffShopBoard

offnews.FPageSize = 1000
offnews.FCurrPage = page
offnews.FScrollCount = 10
offnews.list "30"

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script language='javascript'>
function  TnSearch(frm){
	if (frm.rectuserid.length<1){
		alert('검색어를 입력하세요.');
		return;
	}
	frm.method="get";
	frm.submit();
}
function NextPage(ipage){
	document.frmSrc.page.value= ipage;
	document.frmSrc.submit();
}

function delitems(upfrm){

	var ret = confirm('선택 아이템을 아카데미로 포토앨범으로 이동하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.id.value = upfrm.id.value + frm.id.value + "," ;
				}
			}
		}
		upfrm.action="delalbum.asp";
		upfrm.submit();

	}
}
</script>
<table width="100%" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="50" align="center">번호</td>
    <td width="100" align="center">체크</td>
    <td align="center">제목</td>
    <td width="50" align="center">작성자</td>
    <td width="100" align="center">작성일</td>
  </tr>
<% for i = 0 to (offnews.FResultCount - 1) %>
<form name="frmBuyPrc_<%=i%>" method="post" action="" >
<input type="hidden" name="id" value="<%= offnews.FItemList(i).Fid %>">
  <tr height="20">
    <td align="center">&nbsp;<%= offnews.FItemList(i).Fid %></td>
    <td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
    <td><img src="<%= offnews.FItemList(i).Ficon %>" border="0" alt=""></td>
	<td><%= offnews.FItemList(i).Ftitle %></td>
    <td align="center">Y</td>
    <td align="center"><%= offnews.FItemList(i).Fuserid %></td>
    <td align="center"><%= FormatDate(offnews.FItemList(i).Fregdate, "0000.00.00") %></td>
  </tr>
</form>
<% next %>
</table>
<form method=post name="changefrm">
<input type="hidden" name="id">
</form>
<a href="javascript:delitems(changefrm);">변경하기</a>
<!-- #include virtual="/lib/db/dbclose.asp" -->