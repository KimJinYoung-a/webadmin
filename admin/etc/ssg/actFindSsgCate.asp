<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Response.CharSet = "euc-kr"


Dim ossg, i, page, srcKwd, isNull4DpethNm
page		= requestCheckVar(request("page"),10)
srcKwd		= Trim(requestCheckVar(request("srcKwd"),50))

If page = ""	Then page = 1
'// 목록 접수
Set ossg = new Cssg
	ossg.FPageSize = 30000
	ossg.FCurrPage = page
	ossg.FsearchName = srcKwd
	ossg.getssgCateList
%>
<script>
function chkThis(comp){
    //AnCheckClick(comp);
}

function fnChkThisCate(ii,stdcate,dispcate){
    var iobj;
    if (document.resultFrm.chk.length){
        iobj = document.resultFrm.chk[ii];
    }else{
        iobj = document.resultFrm.chk
    }
    var pchecked = iobj.checked;
    iobj.checked = !pchecked;

}
</script>
<p>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="right"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=ossg.FtotalCount%></strong>&nbsp;&nbsp;</td>
</tr>
</table>
<form name="resultFrm" >
<input type="hidden" name="cdl" value="">
<input type="hidden" name="cdm" value="">
<input type="hidden" name="cds" value="">
<input type="hidden" name="mode" value="saveCateArr">

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
    <td></td>
	<td>DepthCode</td>
	<td>전시매장</td>
	<td>관리카테고리</td>
	<td>Depth1Name</td>
	<td>Depth2Name</td>
	<td>Depth3Name</td>
	<td>Depth4Name</td>
	<td>어린이</td>
	<td>안전</td>
	<td>전기</td>
	<td>위해</td>
</tr>
<% If ossg.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to ossg.FresultCount - 1
			If Trim(ossg.FItemList(i).Fdepth4Nm) = "" Then
				isNull4DpethNm = ossg.FItemList(i).Fdepth3Nm
			Else
				isNull4DpethNm = ossg.FItemList(i).Fdepth4Nm
			End If
%>
<tr align="center" height="25"  title="카테고리 선택" bgcolor="#FFFFFF">
	<td>
	    <input type="checkbox" name="chk" id="chk" value="<%=i%>" onClcik="chkThis(this)";>
	    <input type="hidden" name="depthcode" value="<%= ossg.FItemList(i).FdepthCode %>">
	    <input type="hidden" name="stdcode" value="<%= ossg.FItemList(i).FStdDepthCode %>">
	    <input type="hidden" name="siteno" value="<%= ossg.FItemList(i).Fsiteno %>">
	</td>
	<td><%= ossg.FItemList(i).FdepthCode %></td>
	<td><%= ossg.FItemList(i).getSiteNoToSiteName %></td>
	<td align="left"><%= ossg.FItemList(i).getMmgCateFullName %></td>
	<td><%= ossg.FItemList(i).Fdepth1Nm %></td>
	<td><%= ossg.FItemList(i).Fdepth2Nm %></td>
	<td><%= ossg.FItemList(i).Fdepth3Nm %></td>
	<td><%= ossg.FItemList(i).Fdepth4Nm %></td>
	<td><%= ossg.FItemList(i).FIsChildrenCate %></td>
	<td><%= ossg.FItemList(i).FIssafeCertTgtYn %></td>
	<td><%= ossg.FItemList(i).FIsElecCate %></td>
	<td><%= ossg.FItemList(i).FIsharmCertTgtYn %></td>
</tr>
<%
		Next
	End If
%>
</table>
</form>
<% Set ossg = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
