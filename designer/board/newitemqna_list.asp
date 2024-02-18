<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->
<%
response.write "사용중지된 (구)어드민 입니다. 리뉴얼된 신버전으로 접속해 주세요."
dbget.Close : response.end

dim notupbea, imageon, mifinish, makerid, research ,page, itemid
mifinish = requestCheckVar(request("mifinish"),10)
makerid = session("ssBctID")
research = requestCheckVar(request("research"),10)
page=requestCheckVar(request("page"),10)
itemid=requestCheckVar(request("itemid"),20)
if page="" then page=1

if research="" and mifinish="" then mifinish="on"

dim itemqna
set itemqna = new CItemQna
itemqna.FPageSize = 20
itemqna.FCurrPage=page
itemqna.FRectMakerid = makerid

if (itemid = "") then
	itemqna.FReckMiFinish = mifinish
end if

itemqna.FRectItemID = itemid

itemqna.ItemQnaList

dim i

%>

<script>
function NextPage(pg){
	document.location.href="/designer/board/newitemqna_list.asp?mifinish=<%=mifinish %>&research=<%=research %>&page="+ pg;
}

function isUInt(val) {
	var re = /^[0-9]+$/;
	return re.test(val);
}

function trimString(val) {
    return val.replace(/^\s+|\s+$/gm,'');
}

function SubmitFrm(frm) {
	frm.itemid.value = trimString(frm.itemid.value);

	if (frm.itemid.value != "") {
		if (isUInt(frm.itemid.value) != true) {
			alert("상품코드는 숫자만 가능합니다.");
			return;
		}
	}

	frm.submit();
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	<input type=checkbox name=mifinish <% if mifinish="on" then response.write "checked" %> >미처리만검색
				&nbsp;
				상품코드 : <input type="text" class="text" name="itemid" size="12" value="<%=itemid%>" >
	        </td>
	        <td align="right">
	        	<a href="javascript:SubmitFrm(document.frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
  <tr bgcolor="<%= adminColor("tabletop") %>">
    <td width="150" align="center">고객명(아이디)</td>
    <td align="center">내용</td>
    <td width="60" align="center">상품ID</td>
    <td width="80" align="center">브랜드</td>
    <td width="60" align="center">배송구분</td>
    <td width="80" align="center">작성일</td>
    <td width="80" align="center">답변자</td>
    <td width="80" align="center">답변일</td>
  </tr>
  <%
  for i = 0 to (itemqna.FResultCount - 1)
	  if IsNull(itemqna.FItemList(i).Ftitle) then
		  itemqna.FItemList(i).Ftitle = ""
	  end if

	  if (itemqna.FItemList(i).Ftitle = "") then
		  itemqna.FItemList(i).Ftitle = "(내용없음)"
	  end if
  %>
  <tr bgcolor="#FFFFFF" >
    <td>&nbsp;<%= itemqna.FItemList(i).Fusername %>(<%= itemqna.FItemList(i).Fuserid %>)</td>
    <td>&nbsp;<a href="newitemqna_view.asp?id=<%= itemqna.FItemList(i).Fid %>&menupos=<%= menupos %>"><%= Replace(db2html(itemqna.FItemList(i).Ftitle), "<", "&lt;") %></a></td>
    <td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FItemList(i).FItemID %>" target=_blank><%= itemqna.FItemList(i).FItemID %></a></td>
    <td align="center"><%= itemqna.FItemList(i).FBrandName %></td>
    <td align="center"><%= itemqna.FItemList(i).GetDeliveryTypeName %></td>
    <td align="center"><%= FormatDate(itemqna.FItemList(i).Fregdate, "0000-00-00") %></td>
    <td align="center"><%= itemqna.FItemList(i).Freplyuser %></td>
    <td align="center">
    <% if Not IsNULL(itemqna.FItemList(i).FReplydate) then %>
    <%= FormatDate(itemqna.FItemList(i).FReplydate, "0000-00-00") %>
    <% end if %>
    </td>
  </tr>
<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if itemqna.HasPreScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage - 1) %>')">[prev]</a>
			<% else %>
				[prev]
			<% end if %>
			<% for i = itemqna.StartScrollPage to (itemqna.StartScrollPage + itemqna.FScrollCount - 1) %>
			  <% if (i > itemqna.FTotalPage) then Exit For %>
			  <% if CStr(i) = CStr(itemqna.FCurrPage) then %>
				 [<%= i %>]
			  <% else %>
				 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
			  <% end if %>
			<% next %>
			<% if itemqna.HasNextScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage + itemqna.FScrollCount) %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
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
set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
