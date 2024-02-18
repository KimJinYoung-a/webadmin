<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/history/cs_memocls.asp" -->
<%

dim i, userid, orderserial, divcd, contents_jupsu, backwindow, id,contents_div , mmGubun, phoneNumber, qadiv
dim mode, sqlStr
dim isEditMode
dim sitename

userid          = RequestCheckVar(request("userid"),32)
orderserial     = RequestCheckVar(request("orderserial"),11)
mode            = RequestCheckVar(request("mode"),32)
contents_jupsu  = request("contents_jupsu")
backwindow      = RequestCheckVar(request("backwindow"),32)
id              = RequestCheckVar(request("id"),9)
contents_div    = RequestCheckVar(request("contents_div"),9)
divcd           = RequestCheckVar(request("divcd"),32)

mmGubun         = RequestCheckVar(request("mmGubun"),32)
phoneNumber     = RequestCheckVar(request("phoneNumber"),16)
qadiv           = RequestCheckVar(request("qadiv"),16)
sitename        = RequestCheckVar(request("sitename"),32)
if contents_jupsu <> "" then
	if checkNotValidHTML(contents_jupsu) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if (backwindow = "") then
        backwindow = "opener"
end if

dim ocsmemo
set ocsmemo = New CCSMemo

if (id <> "") then
	ocsmemo.FRectId = id
	ocsmemo.FRectUserID = userid
	ocsmemo.FRectOrderserial = orderserial
	ocsmemo.GetCSMemoDetail

	userid = ocsmemo.FOneItem.FUserID
	orderserial = ocsmemo.FOneItem.Forderserial
	phoneNumber = ocsmemo.FOneItem.FphoneNumber
	sitename = ocsmemo.FOneItem.Fsitename
else
	ocsmemo.GetCSMemoBlankDetail
end if

isEditMode = (id <> "")

'==============================================================================
if (mode = "write") then
	'신규저장모드
    if (divcd = "1") then
		'단순메모
        sqlStr = " insert into " & TABLE_CS_MEMO & "(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,sitename, finishdate,regdate) "
        sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','Y', '" & sitename & "',getdate(),getdate()) "

        dbget.Execute sqlStr
    else
		'요청메모
        sqlStr = " insert into " & TABLE_CS_MEMO & "(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, contents_jupsu, finishyn,sitename,regdate) "
        sqlStr = sqlStr + " values('" + CStr(orderserial) + "','2','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','N', '" & sitename & "',getdate()) "

        dbget.Execute sqlStr
    end if

    response.write "<script>alert('등록되었습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
    dbget.close()	:	response.End
elseif (mode = "modify") then		'수정모드
        sqlStr = " update " & TABLE_CS_MEMO & " "
        sqlStr = sqlStr + " set divcd = '" + CStr(divcd) + "'"
        sqlStr = sqlStr + " , mmgubun = '" + CStr(mmgubun) + "'"
        sqlStr = sqlStr + " , qadiv = '" + CStr(qadiv) + "'"
        sqlStr = sqlStr + " , contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        dbget.Execute sqlStr
		'response.write sqlStr&"<br>"
        response.write "<script>alert('수정되었습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
elseif (mode = "finish") then
        sqlStr = " update " & TABLE_CS_MEMO & " "
        sqlStr = sqlStr + " set finishyn = 'Y'"
        sqlStr = sqlStr + " , finishuser = '" + session("ssBctId") + "'"
        sqlStr = sqlStr + " , finishdate = getdate() "
        sqlStr = sqlStr + " , mmgubun = '" + CStr(mmgubun) + "'"
        sqlStr = sqlStr + " , qadiv = '" + CStr(qadiv) + "'"
        sqlStr = sqlStr + " , contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
        sqlStr = sqlStr + " where id = '" &id&"'"
        'response.write sqlstr
        dbget.Execute sqlStr

        response.write "<script>alert('완료되었습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
elseif (mode = "delete") then
        sqlStr = " delete from " & TABLE_CS_MEMO & " " + VbCrlf
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        dbget.Execute sqlStr

        response.write "<script>alert('삭제되었습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
end if


'=============================================================================
%>
<script>

function GotoHistoryMemoMidify(id,userid,orderserial)
{
frm.action="/cscenter/history/history_memo_write.asp?id=" + id + "&backwindow=" + "opener" + "&userid=" + userid + "&orderserial=" + orderserial
frm.submit();
}
function SubmitForm()
{
        alert("a");
}

function SubmitSave()
{
    if ((document.frm.orderserial.value.length<1)&&(document.frm.userid.value.length<1)&&(document.frm.phoneNumber.value.length<1)) {
	    alert("전화번호, 주문번호, 아이디 중 하나는 입력 되어야 합니다.");
		return;
	}

	if (document.frm.contents_jupsu.value == "") {
		alert("메모내용을 입력하세요.");
		document.frm.contents_jupsu.focus();
		return;
	}

	if (document.frm.qadiv.value.length<1){
	    alert("문의 유형을 선택 하세요.");
		document.frm.qadiv.focus();
		return;
	}

	if(document.frm.id.value == "") {
	    document.frm.mode.value = "write";
	    document.frm.submit();
	}else{
	    document.frm.mode.value = "modify";
	    document.frm.submit();
	}
}

function SubmitFinish()
{
	if (document.frm.contents_jupsu.value == "") {
				alert("메모내용을 입력하세요.");
				return;
				}
        if (confirm("완료처리하겠습니까?") == true) {
                document.frm.mode.value = "finish";
                document.frm.submit();
        }
}

function SubmitDelete()
{
        if (confirm("삭제하겠습니까?") == true) {
                document.frm.mode.value = "delete";
                document.frm.submit();
        }
}


</script>
<body topmargin=10 leftmargin=10 marginwidth=0 marginheight=0>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS메모작성</b>
        </td>
        <td align="right">
            <input type="button" class="button" value="<%= chkIIF(isEditMode,"수정","저장") %>" onclick="javascript:SubmitSave();">
	       	<input type="button" class="button" value="완료" <%= chkIIF((Not isEditMode) or (ocsmemo.FOneItem.Fdivcd<>"2"),"disabled","") %> onclick="javascript:SubmitFinish();">
	        <input type="button" class="button" value="삭제" <%= chkIIF(isEditMode,"","disabled") %> onclick="javascript:SubmitDelete();">
	        <input type="button" class="button" value="닫기" onclick="javascript:window.close();">
	    </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" method="post">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="id" value="<%= ocsmemo.FOneItem.Fid %>">
    <tr height="25">
        <td width="40" bgcolor="<%= adminColor("tabletop") %>">사이트</td>
    	<td bgcolor="#FFFFFF"><%= sitename %></td>
    </tr>
	<tr>
        <td width="40" bgcolor="<%= adminColor("tabletop") %>">전화<br>번호</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="phoneNumber" class="text_ro" value="<%= phoneNumber %>" size="30" readonly></td>
    </tr>
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="orderserial" class="text_ro" value="<%= orderserial %>" size="30" readonly></td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">고객ID</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="userid" class="text_ro" value="<%= userid %>" size="30" readonly></td>
    </tr>
    <% if id = "" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">접수일</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.fregdate %>" size="30" readonly>&nbsp;
	    		당담자ID : <%= ocsmemo.FOneItem.Fwriteuser %>
	    	</td>
	    </tr>
	<% end if %>
	<% if ucase(ocsmemo.FOneItem.Ffinishyn) <> "Y" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">완료일</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.Ffinishdate %>" size="30" readonly>&nbsp;
	    		당담자ID : <%= ocsmemo.FOneItem.Ffinishuser %>
	    	</td>
	    </tr>
	<% end if %>
	<tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">구분</td>
    	<td bgcolor="#FFFFFF">
    	    처리요청 :
    	    <select class="select" name="divcd" <%= ChkIIF(ocsmemo.FOneItem.Fdivcd<>"","disabled","") %> >
	            <option value="1" <% if ocsmemo.FOneItem.Fdivcd = "1" then %>selected<% end if %>>단순메모</option>
	            <option value="2" <% if ocsmemo.FOneItem.Fdivcd = "2" then %>selected<% end if %>>요청메모</option>
	        </select>

			<!-- #include virtual="/cscenter/memo/mmgubunselectbox.asp"-->
	    </td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">메모내용</td>
    	<td bgcolor="#FFFFFF"><textarea name="contents_jupsu" class="textarea" cols="80" rows="7"><%= db2html(ocsmemo.FOneItem.Fcontents_jupsu) %></textarea></td>
    </tr>

</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->

<script language='javascript'>
function getOnLoad(){
	// /cscenter/memo/mmgubunselectbox.asp 참조
	startRequest('mmgubun','0','');
}

window.onload = getOnLoad;
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
