<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/videoInfoCls.asp"-->
<%
dim i,page

dim frmName, compName, videoDiv

frmName= request("frmName")
compName= request("compName")
videoDiv= request("vDiv")

page        = request("page")
if page="" then page=1

dim oVideo
set oVideo = New CVideo
oVideo.FCurrPage = page
oVideo.FPageSize=40
oVideo.FRectDiv = videoDiv
oVideo.FRectUsing = "Y"
oVideo.GetVideoList
%>

<script language='javascript'>
function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

function selectThis(selval){
    opener.<%= frmName %>.<%= compName %>.value = selval;
    window.close();
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="frmName" value="<%= frmName %>">
	<input type="hidden" name="compName" value="<%= compName %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
				동영상 구분 : <%=drawVDivSelect("vDiv",videoDiv)%>
	        </td>
	        <td align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="#BABABA"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
       	<td align="right">
       		총 <%= oVideo.FtotalCount %>건
       	</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>번호</td>
		<td>구분</td>
		<td>썸네일</td>
		<td>제목</td>
		<td>너비</td>
		<td>높이</td>
		<td>등록일</td>
	</tr>
	<% for i=0 to oVideo.FresultCount-1 %>
	<% if oVideo.FItemList(i).Fisusing="Y"	then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
	<tr bgcolor="#EEEEEE">
	<% end if %>
		<td align="center"><a href="javascript:selectThis('<%= oVideo.FItemList(i).FvideoSn %>')"><%= oVideo.FItemList(i).FvideoSn %></a></td>
		<td align="center"><%= getVDivName(oVideo.FItemList(i).FvideoDiv) %></td>
		<td align="center"><a href="javascript:selectThis('<%= oVideo.FItemList(i).FvideoSn %>')"><img src="<%= webImgUrl & "/video/" & oVideo.FItemList(i).FvideoThumb %>" height="50" border="0"></a></td>
		<td align="center"><a href="javascript:selectThis('<%= oVideo.FItemList(i).FvideoSn %>')"><%= oVideo.FItemList(i).FvideoTitle %></a></td>
		<td align="center"><%= oVideo.FItemList(i).FvideoWidth %>px</td>
		<td align="center"><%= oVideo.FItemList(i).FvideoHeight %>px</td>
		<td align="center"><%= left(oVideo.FItemList(i).Fregdate,10) %></td>
	</tr>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if oVideo.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oVideo.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oVideo.StartScrollPage to oVideo.FScrollCount + oVideo.StartScrollPage - 1 %>
    			<% if i>oVideo.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oVideo.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>

        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set oVideo = Nothing
%>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->