<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/callStatic/classes/callstaticCls.asp"-->

<%
	Dim vSessionID, vUserID, vSDate, vEdate, page, i, vMenuPos, vInOut, vDisposition
	vSessionID 	= session("ssBctId")
	vMenuPos	= Request("menupos")
	vUserID		= Request("tenUserID")
	page    	= Request("page")
	vSDate		= Request("sdate")
	vEdate		= Request("edate")
	vInOut		= Request("inout")
	vDisposition= Request("disposition")
	
	If vSDate = "" Then
		vSDate = Left(DateAdd("d",-1,now()),10)
	End If
	If vEdate = "" Then
		vEdate = Left(now(),10)
	End If
	If vInOut = "" Then
		vInOut = "all"
	End If
	If vDisposition = "" Then
		vDisposition = "ANSWERED"
	End If
	
	if page = "" then page = 1 End If
	
	If vUserID = "" Then
		Response.Write "<script>alert('잘못된 접근입니다.');window.close();</script>"
		Response.End
	End IF

	Dim cCallList
	Set cCallList = new ClsCall
	cCallList.FPageSize = 20
	cCallList.FCurrPage = page
	cCallList.FUserID = vUserID
	cCallList.FSDate = vSDate
	cCallList.FEDate = vEdate
	cCallList.FInOut = vInOut
	cCallList.FDisposi = vDisposition
	cCallList.FUserCallList
%>

<script language="Javascript">
function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);
	
	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}
function jsWavPlay(a,b,c)
{
	var winWavPlay;
	winWavPlay = window.open('wav_player.asp?tenUserID='+a+'&yyyymmdd='+b+'&calldate='+c+'','WavPlay','width=300, height=200');
	winWavPlay.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%=vMenuPos%>">
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td rowspan="2" width="100" bgcolor="<%= adminColor("gray") %>">총 건수 : <%=cCallList.ftotalcount%></td>
	<td align="left">
        날짜 : <input type="text" size="10" name="sDate" value="<%=vSDate%>" onClick="jsPopCal('frm','sDate');" style="cursor:hand;">
        ~<input type="text" size="10" name="eDate" value="<%=vEdate%>" onClick="jsPopCal('frm','eDate');" style="cursor:hand;">
		&nbsp;
		상담원ID
		<input type="text" class="text" name="tenUserID" value="<%=vUserID%>" size="12">
		&nbsp;
		발·수신 : <label id="inout1" style="cursor:pointer"><input type="radio" name="inout" id="inout1" value="all" <% If vInOut = "all" Then Response.Write "checked" End IF %>>모두</label>
		<label id="inout2" style="cursor:pointer"><input type="radio" name="inout" id="inout2" value="out" <% If vInOut = "out" Then Response.Write "checked" End IF %>>발신</label>
		<label id="inout3" style="cursor:pointer"><input type="radio" name="inout" id="inout3" value="in" <% If vInOut = "in" Then Response.Write "checked" End IF %>>수신</label>
		&nbsp;
		<select name="disposition">
			<option value="all" <% If vDisposition = "all" Then Response.Write "selected" End If %>>전체</option>
			<option value="ANSWERED" <% If vDisposition = "ANSWERED" Then Response.Write "selected" End If %>>ANSWERED</option>
			<option value="NO ANSWER" <% If vDisposition = "NO ANSWER" Then Response.Write "selected" End If %>>NO ANSWER</option>
			<option value="BUSY" <% If vDisposition = "BUSY" Then Response.Write "selected" End If %>>BUSY</option>
			<option value="FAILED" <% If vDisposition = "FAILED" Then Response.Write "selected" End If %>>FAILED</option>
		</select>
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>날짜</td>
	<td>아이디</td>
	<td>내선</td>
	<td>통화시각</td>
	<td>통화시간</td>
	<td>발·수신번호</td>
	<td>disposition</td>
	<td></td>
</tr>
<% if cCallList.FResultCount > 0 then %>
<% for i=0 to cCallList.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td align="center"><%=cCallList.FItemList(i).fdate%></td>
	<td align="center"><%=cCallList.FItemList(i).fuserid%></td>
	<td align="center"><%=cCallList.FItemList(i).fcomtelno%></td>
	<td align="center"><%=cCallList.FItemList(i).fteltime%></td>
	<td align="center"><%=sec2time(cCallList.FItemList(i).ftelterm)%></td>
	<td align="center"><%=cCallList.FItemList(i).fclienttelno%></td>
	<td align="center"><%=cCallList.FItemList(i).fdisposition%></td>
	<td align="center">
	<%
			If cCallList.FItemList(i).fwavlink <> "x" AND cCallList.FItemList(i).fdisposition = "ANSWERED" Then
				If cCallList.FItemList(i).fuserid = vSessionID OR C_ADMIN_AUTH or C_CSPowerUser Then
	%>
				<input type="button" value="듣기" onClick="jsWavPlay('<%=cCallList.FItemList(i).fuserid%>','<%=cCallList.FItemList(i).fdate%>','<%=cCallList.FItemList(i).fcstrcalldate%>')" class="button">		
	<%
				End If
	 		End If
	%>
	</td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	   	<% if cCallList.HasPreScroll then %>
			<span class="list_link"><a href="?menupos=<%=vMenuPos%>&page=<%= cCallList.StartScrollPage-1 %>&tenUserID=<%=vUserID%>&sdate=<%=vSDate%>&edate=<%=vEdate%>&inout=<%=vInOut%>&disposition=<%=vDisposition%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + cCallList.StartScrollPage to cCallList.StartScrollPage + cCallList.FScrollCount - 1 %>
			<% if (i > cCallList.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(cCallList.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?menupos=<%=vMenuPos%>&page=<%= i %>&tenUserID=<%=vUserID%>&sdate=<%=vSDate%>&edate=<%=vEdate%>&inout=<%=vInOut%>&disposition=<%=vDisposition%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if cCallList.HasNextScroll then %>
			<span class="list_link"><a href="?menupos=<%=vMenuPos%>&page=<%= i %>&tenUserID=<%=vUserID%>&sdate=<%=vSDate%>&edate=<%=vEdate%>&inout=<%=vInOut%>&disposition=<%=vDisposition%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

	
	
<%
	set cCallList = nothing
'DATAMART>>상담원 콜센터 통계  에서
'클릭시 => 팝업, 상세내역 리스트 / 녹취파일 듣기
'팝업창 검색조건 : 날짜, 상담원아이디, 수신, 발신
'리스트 필드
'아이디, 내선, 통화시각, 통화시간, (발)수신번호, disposition, 녹취링크 등.
'
'녹취파일은 본인 또는 coolhas , beso 아이디만 청취가능
'참조 table
'73서버의 db_datamart.dbo.tbl_call_cdr
'참조메뉴
'DATAMART>>상담원 콜센터 통계 및
'DATAMART>>부재중 전화 통계
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
