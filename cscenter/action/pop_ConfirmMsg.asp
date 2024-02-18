<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->

<%

dim id, fin
dim sitegubun

id = request("id")
fin = request("fin")
sitegubun      	= RequestCheckVar(request("sitegubun"),32)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
	if (sitegubun = "academy") then
		ocsaslist.GetOneCSASMasterAcademy
	else
		'10x10
		ocsaslist.GetOneCSASMaster
	end if
end if


''확인요청정보 :
dim OCsConfirm
set OCsConfirm = new CCSASList
OCsConfirm.FRectCsAsID = id

if id<>"" then
	if (sitegubun = "academy") then
		OCsConfirm.GetOneCsConfirmItemAcademy
	else
		'10x10
		OCsConfirm.GetOneCsConfirmItem
	end if
end if


if (ocsaslist.FResultCount<1) then
    response.write "<script>alert('유효하지 않은 내역입니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

''확인요청내역이 없는 경우=> 접수상태만 등록 가능.
''확인요청내역이 있는 경우=> 확인 요청 상태에서만 수정/완료 가능.
if (OCsConfirm.FResultCount<1) then
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('완료 이전 상태 에서만 등록 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
else
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('완료 이전 상태 에서만 수정/완료 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
end if


dim IsEditMode
IsEditMode = (OCsConfirm.FResultCount>0)

%>
<script language='javascript'>
function ActConfirmReg(frm){
    if (frm.confirmregmsg.value.length<1){
        alert('확인요청 내용을 입력해 주세요.');
        frm.confirmregmsg.focus();
        return;
    }

    if (confirm('확인 요청 내용 입력시 상태가 확인 요청중으로 변경되며, 환불파일에서 삭제됩니다. 등록하시겠습니까?')){
        frm.nextstate.value = "B005";
        frm.mode.value = "reg";
        frm.submit();
    }
}

function ActConfirmReEdit(frm){
    if (frm.confirmregmsg.value.length<1){
        alert('확인요청 내용을 입력해 주세요.');
        frm.confirmregmsg.focus();
        return;
    }

    if (confirm('재요청 하시겠습니까?\n 기존 확인 내역은 초기 화 됩니다.')){
        frm.nextstate.value = "B005";
        frm.mode.value = "reInput";
        frm.submit();
    }
}


function ActConfirmEdit(frm){
    if (frm.confirmregmsg.value.length<1){
        alert('확인요청 내용을 입력해 주세요.');
        frm.confirmregmsg.focus();
        return;
    }

    if (confirm('접수된 요청 내용을 수정 하시겠습니까?')){
        frm.mode.value = "edit";
        frm.submit();
    }
}

function ActConfirmFinish(frm){
    if (frm.confirmfinishmsg.value.length<1){
        alert('확인처리 내용을 입력해 주세요.');
        frm.confirmfinishmsg.focus();
        return;
    }

    if (confirm('확인 요청 내용 처리시 상태가 접수상태로 재 변경됩니다. 완료 처리 하시겠습니까?')){
        frm.nextstate.value = "B001";
        frm.mode.value = "finish";
        frm.submit();
    }
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmConfirm" method="post" action="pop_ConfirmMsg_process.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="sitegubun" value="<%= sitegubun %>">
<input type="hidden" name="nextstate" value="">
<input type="hidden" name="mode" value="">

<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">사이트</td>
    <td bgcolor="#FFFFFF">
        <%= sitegubun %>
    </td>
</tr>

<% if IsEditMode then %>

<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">확인요청 등록자</td>
    <td bgcolor="#FFFFFF">
        <%= OCsConfirm.FOneItem.Fconfirmreguserid %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">확인요청내용</td>
    <td bgcolor="#FFFFFF">
        <textarea class="textarea" name="confirmregmsg" cols="48" rows="5"  ><%= OCsConfirm.FOneItem.Fconfirmregmsg %></textarea>
    </td>
</tr>
<% if fin<>"" then %>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">확인요청 처리자</td>
    <td bgcolor="#FFFFFF">
        <%= session("ssBctID") %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">확인처리내용</td>
    <td bgcolor="#FFFFFF">
        <textarea class="textarea" name="confirmfinishmsg" cols="48" rows="5"  ><%= OCsConfirm.FOneItem.Fconfirmfinishmsg %></textarea>
    </td>
</tr>
<% end if %>
<% else %>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">확인요청등록자</td>
    <td bgcolor="#FFFFFF">
        <%= session("ssBctID") %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">확인요청내용</td>
    <td bgcolor="#FFFFFF">
        <textarea class="textarea" name="confirmregmsg" cols="48" rows="5"  ></textarea>
    </td>
</tr>
<% end if %>
<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <% if IsEditMode then %>
        <% if Not IsNULL(OCsConfirm.FOneItem.Fconfirmfinishdate) then %>
        처리 완료된 내역입니다. <br>: (재확인요청시 기존처리내역은 삭제됩니다.)
        <br>
        <% if (fin<>"fin") then %>
        <!-- 재확인요청 변경메뉴 추가 요망 -->
        <input type="button" class="button" value=" 재 확인요청  " onClick="ActConfirmReEdit(frmConfirm)">
        <% end if %>
        <% else %>
            <% if fin="" then %>
            <input type="button" class="button" value=" 확인요청 수정 " onClick="ActConfirmEdit(frmConfirm)">
            <% else %>
            <input type="button" class="button" value=" 확인요청 완료 " onClick="ActConfirmFinish(frmConfirm)">
            <% end if %>
        <% end if %>
    <% else %>
    <input type="button" class="button" value=" 확인요청 등록 " onClick="ActConfirmReg(frmConfirm)">

    <% end if %>
    </td>
</tr>
</form>
</table>
<%
set ocsaslist = Nothing
set OCsConfirm = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->