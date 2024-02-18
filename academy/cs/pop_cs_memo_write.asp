<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%

dim i, userid, orderserial, divcd, contents_jupsu, backwindow, id
dim mode, sqlStr

userid = RequestCheckvar(request("userid"),32)
orderserial = RequestCheckvar(request("orderserial"),16)
mode = RequestCheckvar(request("mode"),16)
divcd = RequestCheckvar(request("divcd"),10)
contents_jupsu = request("contents_jupsu")
backwindow = RequestCheckvar(request("backwindow"),10)
id = RequestCheckvar(request("id"),10)

if ((userid = "") and (orderserial = "") and (id = "")) then
        response.write "<script>alert('잘못된 접속입니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (backwindow = "") then
        backwindow = "opener"
end if


'==============================================================================
if (mode = "write") then
        if (divcd = "1") then
                sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
                sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','" + session("ssBctId") + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','Y',getdate(),getdate()) "
                rsget.Open sqlStr,dbget,1
        else
                sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, writeuser, contents_jupsu, finishyn,regdate) "
                sqlStr = sqlStr + " values('" + CStr(orderserial) + "','2','" + CStr(userid) + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','N',getdate()) "
                rsget.Open sqlStr,dbget,1
        end if

        response.write "<script>alert('등록되었습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
elseif (mode = "modify") then
        sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
        sqlStr = sqlStr + " set divcd = '" + CStr(divcd) + "', contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        rsget.Open sqlStr,dbget,1

        response.write "<script>alert('수정되었습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
elseif (mode = "finish") then
        sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
        sqlStr = sqlStr + " set finishyn = 'Y', finishuser = '" + session("ssBctId") + "',finishdate = getdate() "
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        rsget.Open sqlStr,dbget,1

        response.write "<script>alert('완료되었습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
elseif (mode = "delete") then
        sqlStr = " delete from [db_cs].[dbo].tbl_cs_memo "
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        rsget.Open sqlStr,dbget,1

        response.write "<script>alert('삭제습니다.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
end if


'==============================================================================
dim ocsmemo
set ocsmemo = New CCSMemo

if (id <> "") then
        ocsmemo.FRectId = id
        ocsmemo.GetCSMemoDetail

        userid = ocsmemo.FOneItem.Fuserid
        orderserial = ocsmemo.FOneItem.Forderserial
else
        ocsmemo.GetCSMemoBlankDetail
end if

%>
<script>
function SubmitForm()
{
        alert("a");
}

function SubmitSave()
{
        if (document.frm.contents_jupsu.value == "") {
                alert("메모내용을 입력하세요.");
                return;
        }
<% if (id = "") then %>
        document.frm.mode.value = "write";
<% else %>
        document.frm.mode.value = "modify";
<% end if %>
        document.frm.submit();
}

function SubmitFinish()
{
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


			<table width="100%" height="35" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<tr height="10" valign="bottom">
				    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td background="/images/tbl_blue_round_02.gif"></td>
				    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
				</tr>
				<tr height="25">
				    <td background="/images/tbl_blue_round_04.gif"></td>
				    <td background="/images/tbl_blue_round_06.gif">
				    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS메모작성</b>
				    </td>
				    <td align="right" background="/images/tbl_blue_round_06.gif">
				        <input type="button" value="저장하기" onclick="javascript:SubmitSave();">
				        <input type="button" value="완료하기" onclick="javascript:SubmitFinish();">
				        <input type="button" value="삭제하기" onclick="javascript:SubmitDelete();">
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
			</table>
			<table width="100%" height="195" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
				<form name="frm" onsubmit="return false;" method="post">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="id" value="<%= ocsmemo.FOneItem.Fid %>">
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td width="100">주문번호</td>
				    <td><input type="text" name="orderserial" value="<%= orderserial %>" style='background-color:#DDDDFF' readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>고객ID</td>
				    <td><input type="text" name="userid" value="<%= userid %>" style='background-color:#DDDDFF' readonly></td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="20">
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td>구분</td>
				    <td>
				      <select name="divcd">
				        <option value="1" <% if (ocsmemo.FOneItem.Fdivcd = "1") then %>selected<% end if %>>단순메모</option>
				        <option value="2" <% if (ocsmemo.FOneItem.Fdivcd = "2") then %>selected<% end if %>>요청메모</option>
				      </select>
				    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr>
					<td background="/images/tbl_blue_round_04.gif"></td>
				    <td valign="top">메모내용</td>
				    <td>
				        <textarea name="contents_jupsu" rows="6" cols="35"><%= ocsmemo.FOneItem.Fcontents_jupsu %></textarea>
                    </td>
				    <td background="/images/tbl_blue_round_05.gif"></td>
				</tr>
				<tr height="10" valign="top">
					<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td background="/images/tbl_blue_round_08.gif"></td>
				    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
				</tr>
				</form>
			</table>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->