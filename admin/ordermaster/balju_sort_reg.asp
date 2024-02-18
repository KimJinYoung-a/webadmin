<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ü� ���ļ�������
' History : 2020.12.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->
<%
dim i, menupos, orack, osort
dim midx, title, comment, isusing, regdate, lastupdate, regadminid, lastadminid, didx, rackcode, sortno
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
    midx = requestCheckVar(getNumeric(request("midx")),10)

Set osort = New CTenBalju
	osort.frectmidx = midx

    If midx <> "" Then
        osort.GetBaljusortview()
    end if

if osort.ftotalcount > 0 then
    title = ReplaceBracket(osort.FOneItem.Ftitle)
    comment = ReplaceBracket(osort.FOneItem.Fcomment)
    isusing = osort.FOneItem.Fisusing
    regdate = osort.FOneItem.Fregdate
    lastupdate = osort.FOneItem.Flastupdate
    regadminid = osort.FOneItem.Fregadminid
    lastadminid = osort.FOneItem.Flastadminid
else
    if isusing="" then isusing="Y"
end if

set osort = nothing
%>

<script type="text/javascript">

function checkform(frm){
    if (frm.title.value==""){
        alert('������ �Է��� �ּ���.');
        frm.title.focus();
    }
    if (frm.isusing.value==""){
        alert('��뿩�θ� �ּ���.');
        frm.title.focus();
    }

    var rackcode = document.getElementsByName("rackcode");
    var layer = document.getElementsByName("layer");
    var sortno = document.getElementsByName("sortno");
    for (var i=0; i < rackcode.length; i++){
        if (rackcode[i].value==""){
            alert('���ڵ带 �Է��ϼ���.');
            rackcode[i].focus();
            return;
        }
        if (layer[i].value==""){
            alert('���� �Է��ϼ���.');
            layer[i].focus();
            return;
        }
        if (sortno[i].value==""){
            alert('���ļ����� �Է��ϼ���.');
            sortno[i].focus();
            return;
        }
    }

	frm.submit();
}

// ���û���
function delSelectdrackcode(){
	if(confirm("������ ���ڵ带 �����Ͻðڽ��ϱ�?"))
		tablerackcode.deleteRow(tablerackcode.clickedRowIndex);
}

//���ڵ� �߰�
function addSelectedrackcode(){
	var lenRow = tablerackcode.rows.length;

	// ���߰�
	var oRow = tablerackcode.insertRow(lenRow);
	oRow.onmouseover=function(){tablerackcode.clickedRowIndex=this.rowIndex};

	// ���߰� (�̸�,������ư)
	var oCell1 = oRow.insertCell(0);		
	var oCell2 = oRow.insertCell(1);

	oCell1.innerHTML = "���ڵ�:<input type='text' name='rackcode' value='' size=30 maxlength=32 > / ��:<input type='text' name='layer' value='1' size=8 maxlength=10 > / ���ļ���:<input type='text' name='sortno' value='0' size=8 maxlength=10 >";
	oCell2.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdrackcode()' align='absmiddle'>";
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
        * �� ���ļ��� ����          
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frm" action="/admin/ordermaster/balju_sort_process.asp" method="post" style="margin:0px;" >
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" name="mode" value="baljusortreg">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<% If midx <> "" Then %>
<tr height="30">
    <td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
    <td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= midx %></td>
</tr>
<% End If %>
<tr height="30">
    <td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
    <td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
        <input type="text" name="title" value="<%= title %>" size=100 maxlength=128 >
    </td>
</tr>
<tr >
    <td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">���ڵ�����</td>
    <td valign="bottom" bgcolor="#FFFFFF" >
        <table border="0" cellspacing="0" class="a">
        <tr id="disprackcodediv" name="disprackcodediv" style="display:">
            <td>
                �������� ���� �켱���� : 1���� ��(��������) , 2���� ���ļ���(��������)
                <table name='tablerackcode' id='tablerackcode' class=a>
                <%
                If midx <> "" Then
                    set orack = new CTenBalju
                        orack.frectmidx = midx
                        orack.GetBaljusortrackcodelist

                    if orack.FResultCount > 0 then
                
                    for i=0 to orack.FResultCount-1
                %>
                    <tr onMouseOver='tablerackcode.clickedRowIndex=this.rowIndex'>
                        <td>
                            ���ڵ�:<input type='text' name='rackcode' value='<%= orack.FItemList(i).frackcode %>' size=30 maxlength=32 > / ��:<input type='text' name='layer' value='<%= orack.FItemList(i).flayer %>' size=8 maxlength=10 > / ���ļ���:<input type='text' name='sortno' value='<%= orack.FItemList(i).fsortno %>' size=8 maxlength=10 >    
                        </td>  
                        <td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdrackcode()' align='absmiddle'></td>   
                    </tr>
                <%
                    next
                    end if
                    set orack=nothing
                end if
                %>
                </table>
            </td>
            <td valign="bottom">
                <input type="button" class="button" value="���ڵ��߰�" onclick="addSelectedrackcode()">
            </td>
        <tr>
        </table>
    </td>
</tr>
<tr height="30">
    <td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�ڸ�Ʈ</td>
    <td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
        <textarea name="comment" id="comment" style="width: 700px; height: 200px;"><%= comment %></textarea>
    </td>
</tr>
<tr height="30">
    <td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
    <td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
        <% drawSelectBoxisusingYN "isusing",isusing,"" %>
    </td>
</tr>
<% If midx <> "" Then %>
<tr height="30">
    <td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�����</td>
    <td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
        <%= regdate %>
        <Br><%= regadminid %>
    </td>
</tr>
<tr height="30">
    <td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����������</td>
    <td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
        <% if lastupdate<>"" then %>
            <%= lastupdate %>
        <% end if %>
        <% if lastadminid<>"" then %>
            <Br><%= lastadminid %>
        <% end if %>
    </td>
</tr>
<% End If %>

<tr height="30">
    <td align="center" bgcolor="#FFFFFF" colspan=2>
        <input type="button" onclick="checkform(frm);" value="�����ϱ�" class="button">
    </td>	
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->