<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/contestCls.asp"-->

<%
	Dim cPollList, vDiv, iTotCnt, i, vUserNum, vUserID, vSubject, vContents
	vDiv = requestCheckVar(Request("divnum"),10)
	vUserNum = Request("usernum")
	
	IF vUserNum <> "" Then
		set cPollList = new ClsContest
		cPollList.FDiv = vDiv
		cPollList.FUserNum = vUserNum
		cPollList.fevt_ContestEdit()
		vUserID		= cPollList.FOneItem.fuserid
		vSubject	= cPollList.FOneItem.fsubject
		vContents	= cPollList.FOneItem.fcontents
		set cPollList = nothing
	End If
	
	set cPollList = new ClsContest
	cPollList.FDiv = vDiv
	cPollList.FFinalList()
	iTotCnt = cPollList.FResultCount
%>

<script language="javascript">
function delproc(idx)
{
	if(confirm("선택하신 아이디를 삭제하시겠습니까?") == true) {
		document.delprocfrm.idx.value = idx;
		document.delprocfrm.submit();
	}
}

function pollplusproc(idx)
{
	document.pollplusfrm.idx.value = idx;
	document.pollplusfrm.submit();
}

function submitfrm()
{
	if(frm1.userid.value == "")
	{
		alert("회원아이디를 입력하세요.");
		frm1.userid.focus();
		return false;
	}
	
	if(confirm("입력한 내용을 저장하시겠습니까?") == true) {
		frm1.submit();
	}
}

function editpoll(usernum)
{
	location.href = "popup_finallist.asp?divnum=<%=vDiv%>&usernum="+usernum+"";
}

function rewrite()
{
	if(confirm("진짜 다시 쓰시겠습니까?") == true) {
		location.href = "popup_finallist.asp?divnum=<%=vDiv%>";
	}
}

function viewtext(i)
{
	if(document.getElementById("text"+i+"").style.display == "none")
	{
		document.getElementById("text"+i+"").style.display = "block";
	}
	else
	{
		document.getElementById("text"+i+"").style.display = "none";
	}
}

function imageupload(unum){
	var pop_image = window.open('popup_finallist_image.asp?contest=<%=vDiv%>&usernum='+unum+'','pop_image','width=400,height=700,scrollbars=yes,resizable=yes');
	pop_image.focus();
}
</script>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" action="popup_finallist_proc.asp" method="post">
<input type="hidden" name="divnum" value="<%=vDiv%>">
<input type="hidden" name="idx" value="<%=vUserNum%>">
<input type="hidden" name="gubun" value="insert">
<tr bgcolor="FFFFFF">
	<td>회원아이디 : <input type="text" name="userid" value="<%=vUserID%>" size="13" tabindex="1" <% If vUserNum <> "" Then Response.Write "readonly" End If %>></td>
	<td>Subject : <input type="text" name="subject" value="<%=vSubject%>" size="50" maxlength="50" tabindex="2"></td>
	<td rowspan="2" align="center"><input type="button" value="저 장" onClick="submitfrm();"><br><br><input type="button" value="다시쓰기" onClick="rewrite();"></td>
</tr>
<tr bgcolor="FFFFFF">
	<td colspan="2">내용: <textarea name="contents" cols="80" rows="3" tabindex="3"><%=vContents%></textarea></td>
</tr>
</form>
</table>
<br>

<table cellpadding="0" cellspacing="0" class="a">
<tr height="25">
	<td colspan="20">
		Total Count : <b><%= iTotCnt %></b>
	</td>
</tr>
<%
	If cPollList.FResultCount <> 0 Then
		For i = 0 To cPollList.FResultCount -1
%>
		<tr bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
			<td>
				<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<% If i = 0 Then %>
				<tr align="center" bgcolor="#E6E6E6" height="20">
					<td>파이널 리스트</td>
					<td>Subject</td>
					<td></td>
					<td></td>
				</tr>
				<% End If %>
				<tr bgcolor="FFFFFF">
					<td width="170" style="cursor:pointer" onClick="viewtext('<%=i%>');"><%=cPollList.FItemList(i).fusername%>(<%=cPollList.FItemList(i).fuserid%>)</td>
					<td width="350" style="cursor:pointer" onClick="viewtext('<%=i%>');"><%=cPollList.FItemList(i).fsubject%></td>
					<td align="center">
					[<a href="javascript:editpoll('<%=cPollList.FItemList(i).fusernum%>');">수정</a>]
					[<a href="javascript:delproc('<%=cPollList.FItemList(i).fusernum%>');">삭제</a>]
					[<a href="javascript:pollplusproc('<%=cPollList.FItemList(i).fusernum%>');">투표+1</a>]
					</td>
					<td><input type="button" value="이미지" onClick="imageupload('<%=cPollList.FItemList(i).fusernum%>')"></td>
				</tr>
				<tr bgcolor="FFFFFF" id="text<%=i%>" style="display:none;">
					<td colspan="4" style="padding:5 3 5 3;">
						<%=Replace(cPollList.FItemList(i).fcontents,vbCrLf,"<br>")%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
</table>
<form name="delprocfrm" action="popup_finallist_proc.asp" method="post">
<input type="hidden" name="gubun" value="del">
<input type="hidden" name="idx" value="">
<input type="hidden" name="divnum" value="<%=vDiv%>">
</form>
<form name="pollplusfrm" action="popup_finallist_proc.asp" method="post">
<input type="hidden" name="gubun" value="pollplus">
<input type="hidden" name="idx" value="">
<input type="hidden" name="divnum" value="<%=vDiv%>">
</form>
<%
	set cPollList = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->