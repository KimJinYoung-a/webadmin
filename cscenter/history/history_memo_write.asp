<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸� 
' History : 2007.10.26 �ѿ�� ����
'###########################################################
%> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->

������� �޴� - ������ ���� ���

<%
dbget.close()	:	response.End

dim i, userid, orderserial, divcd, contents_jupsu, backwindow, id,contents_div , mmGubun, phoneNumber, qadiv
dim mode, sqlStr 
dim isEditMode

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
else
	ocsmemo.GetCSMemoBlankDetail
end if

isEditMode = (id <> "")

'==============================================================================
if (mode = "write") then	'�ű�������
        if (divcd = "1") then		'�ܼ��޸�
                sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
                sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','Y',getdate(),getdate()) "
            
                dbget.Execute sqlStr
        else			'��û�޸�
                sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, contents_jupsu, finishyn,regdate) "
                sqlStr = sqlStr + " values('" + CStr(orderserial) + "','2','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','N',getdate()) "
                
                dbget.Execute sqlStr
        end if

        response.write "<script>alert('��ϵǾ����ϴ�.'); " + backwindow + ".location.reload(); " + backwindow + ".focus(); window.close();</script>"
        dbget.close()	:	response.End
elseif (mode = "modify") then		'�������
        sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
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
        sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
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
        sqlStr = " delete from [db_cs].[dbo].tbl_cs_memo " + VbCrlf
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
    	    <select name="divcd" <%= ChkIIF(ocsmemo.FOneItem.Fdivcd<>"","disabled","") %> >
	            <option value="1" <% if ocsmemo.FOneItem.Fdivcd = "1" then %>selected<% end if %>>�ܼ��޸�</option>
	            <option value="2" <% if ocsmemo.FOneItem.Fdivcd = "2" then %>selected<% end if %>>��û�޸�</option>
	        </select>
	        
	        �޸𱸺�
	        <select name="mmGubun">
	            <option value="0" <% if ocsmemo.FOneItem.FmmGubun = "0" then %>selected<% end if %>>�Ϲݸ޸�</option>
	            <option value="1" <% if ocsmemo.FOneItem.FmmGubun = "1" then %>selected<% end if %>>�ιٿ����ȭ</option>
	            <option value="2" <% if ocsmemo.FOneItem.FmmGubun = "2" then %>selected<% end if %>>�ƿ��ٿ����ȭ</option>
	            <option value="3" <% if ocsmemo.FOneItem.FmmGubun = "3" then %>selected<% end if %>>��ü��ȭ</option>
	            <!--
	            <option value="4" <% if ocsmemo.FOneItem.FmmGubun = "4" then %>selected<% end if %>>SMS</option>
	            <option value="5" <% if ocsmemo.FOneItem.FmmGubun = "5" then %>selected<% end if %>>EMAIL</option>
	            -->
	        </select>
	        
	        ���� :
  			<select class="select" name="qadiv">
                <option value="">��ü</option>
                <option value="00" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="00","selected","") %> >��۹���</option>
                <option value="01" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="01","selected","") %> >�ֹ�����</option>
                <option value="02" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="02","selected","") %> >��ǰ����</option>
                <option value="03" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="03","selected","") %> >�����</option>
                <option value="04" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="04","selected","") %> >��ҹ���</option>
                <option value="05" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="05","selected","") %> >ȯ�ҹ���</option>
                <option value="06" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="06","selected","") %> >��ȯ����</option>
                <option value="07" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="07","selected","") %> >AS����</option>    
                <option value="08" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="08","selected","") %> >�̺�Ʈ����</option>
                <option value="09" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="09","selected","") %> >������������</option>    
                <option value="10" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="10","selected","") %> >�ý��۹���</option>
                <option value="11" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="11","selected","") %> >ȸ����������</option>
                <option value="12" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="12","selected","") %> >ȸ����������</option>
                <option value="13" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="13","selected","") %> >��÷����</option>
                <option value="14" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="14","selected","") %> >��ǰ����</option>
                <option value="15" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="15","selected","") %> >�Աݹ���</option>
                <option value="16" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="16","selected","") %> >�������ι���</option>
                <option value="17" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="17","selected","") %> >����/���ϸ�������</option>
                <option value="18" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="18","selected","") %> >�����������</option>
                <option value="20" <%= ChkIIF(ocsmemo.FOneItem.Fqadiv="20","selected","") %> >��Ÿ����</option>
            </select>
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


<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->







