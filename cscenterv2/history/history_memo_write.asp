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
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
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
	'�ű�������
    if (divcd = "1") then
		'�ܼ��޸�
        sqlStr = " insert into " & TABLE_CS_MEMO & "(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,sitename, finishdate,regdate) "
        sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','Y', '" & sitename & "',getdate(),getdate()) "

        dbget.Execute sqlStr
    else
		'��û�޸�
        sqlStr = " insert into " & TABLE_CS_MEMO & "(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, contents_jupsu, finishyn,sitename,regdate) "
        sqlStr = sqlStr + " values('" + CStr(orderserial) + "','2','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','N', '" & sitename & "',getdate()) "

        dbget.Execute sqlStr
    end if

    response.write "<script>alert('��ϵǾ����ϴ�.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
    dbget.close()	:	response.End
elseif (mode = "modify") then		'�������
        sqlStr = " update " & TABLE_CS_MEMO & " "
        sqlStr = sqlStr + " set divcd = '" + CStr(divcd) + "'"
        sqlStr = sqlStr + " , mmgubun = '" + CStr(mmgubun) + "'"
        sqlStr = sqlStr + " , qadiv = '" + CStr(qadiv) + "'"
        sqlStr = sqlStr + " , contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        dbget.Execute sqlStr
		'response.write sqlStr&"<br>"
        response.write "<script>alert('�����Ǿ����ϴ�.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
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

        response.write "<script>alert('�Ϸ�Ǿ����ϴ�.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
elseif (mode = "delete") then
        sqlStr = " delete from " & TABLE_CS_MEMO & " " + VbCrlf
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        dbget.Execute sqlStr

        response.write "<script>alert('�����Ǿ����ϴ�.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
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
	    alert("��ȭ��ȣ, �ֹ���ȣ, ���̵� �� �ϳ��� �Է� �Ǿ�� �մϴ�.");
		return;
	}

	if (document.frm.contents_jupsu.value == "") {
		alert("�޸𳻿��� �Է��ϼ���.");
		document.frm.contents_jupsu.focus();
		return;
	}

	if (document.frm.qadiv.value.length<1){
	    alert("���� ������ ���� �ϼ���.");
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
				alert("�޸𳻿��� �Է��ϼ���.");
				return;
				}
        if (confirm("�Ϸ�ó���ϰڽ��ϱ�?") == true) {
                document.frm.mode.value = "finish";
                document.frm.submit();
        }
}

function SubmitDelete()
{
        if (confirm("�����ϰڽ��ϱ�?") == true) {
                document.frm.mode.value = "delete";
                document.frm.submit();
        }
}


</script>
<body topmargin=10 leftmargin=10 marginwidth=0 marginheight=0>

<!-- ǥ ��ܹ� ����-->
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
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS�޸��ۼ�</b>
        </td>
        <td align="right">
            <input type="button" class="button" value="<%= chkIIF(isEditMode,"����","����") %>" onclick="javascript:SubmitSave();">
	       	<input type="button" class="button" value="�Ϸ�" <%= chkIIF((Not isEditMode) or (ocsmemo.FOneItem.Fdivcd<>"2"),"disabled","") %> onclick="javascript:SubmitFinish();">
	        <input type="button" class="button" value="����" <%= chkIIF(isEditMode,"","disabled") %> onclick="javascript:SubmitDelete();">
	        <input type="button" class="button" value="�ݱ�" onclick="javascript:window.close();">
	    </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" method="post">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="id" value="<%= ocsmemo.FOneItem.Fid %>">
    <tr height="25">
        <td width="40" bgcolor="<%= adminColor("tabletop") %>">����Ʈ</td>
    	<td bgcolor="#FFFFFF"><%= sitename %></td>
    </tr>
	<tr>
        <td width="40" bgcolor="<%= adminColor("tabletop") %>">��ȭ<br>��ȣ</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="phoneNumber" class="text_ro" value="<%= phoneNumber %>" size="30" readonly></td>
    </tr>
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="orderserial" class="text_ro" value="<%= orderserial %>" size="30" readonly></td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">��ID</td>
    	<td bgcolor="#FFFFFF"><input type="text" name="userid" class="text_ro" value="<%= userid %>" size="30" readonly></td>
    </tr>
    <% if id = "" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">������</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.fregdate %>" size="30" readonly>&nbsp;
	    		�����ID : <%= ocsmemo.FOneItem.Fwriteuser %>
	    	</td>
	    </tr>
	<% end if %>
	<% if ucase(ocsmemo.FOneItem.Ffinishyn) <> "Y" then %>
    <% else %>
	    <tr>
	    	<td bgcolor="<%= adminColor("tabletop") %>">�Ϸ���</td>
	    	<td bgcolor="#FFFFFF">
	    		<input type="text" name="regdate" class="text_ro" value="<%= ocsmemo.FOneItem.Ffinishdate %>" size="30" readonly>&nbsp;
	    		�����ID : <%= ocsmemo.FOneItem.Ffinishuser %>
	    	</td>
	    </tr>
	<% end if %>
	<tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF">
    	    ó����û :
    	    <select class="select" name="divcd" <%= ChkIIF(ocsmemo.FOneItem.Fdivcd<>"","disabled","") %> >
	            <option value="1" <% if ocsmemo.FOneItem.Fdivcd = "1" then %>selected<% end if %>>�ܼ��޸�</option>
	            <option value="2" <% if ocsmemo.FOneItem.Fdivcd = "2" then %>selected<% end if %>>��û�޸�</option>
	        </select>

			<!-- #include virtual="/cscenter/memo/mmgubunselectbox.asp"-->
	    </td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">�޸𳻿�</td>
    	<td bgcolor="#FFFFFF"><textarea name="contents_jupsu" class="textarea" cols="80" rows="7"><%= db2html(ocsmemo.FOneItem.Fcontents_jupsu) %></textarea></td>
    </tr>

</table>

<!-- ǥ �ϴܹ� ����-->
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
<!-- ǥ �ϴܹ� ��-->

<script language='javascript'>
function getOnLoad(){
	// /cscenter/memo/mmgubunselectbox.asp ����
	startRequest('mmgubun','0','');
}

window.onload = getOnLoad;
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->
