<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/topbrand/topbrandtalkcls.asp" -->
<%

dim i, page, iscurrtopbrand, idx

page = requestCheckVar(request("page"),10)
if (page = "") then
    page = "1"
end if

idx = requestCheckVar(request("idx"),20)

'==============================================================================
dim otopbrandtalklist
set otopbrandtalklist = New CTopBrandTalk

iscurrtopbrand = otopbrandtalklist.IsCurrentTopBrand(session("ssBctId"))

otopbrandtalklist.FRectMakerID = session("ssBctId")
otopbrandtalklist.FRectIdx = idx
'otopbrandtalklist.FRectIsCurrentTopBrand = "Y"

otopbrandtalklist.GetTopBrandTalkOne

%>
<script>
function SubmitWrite()
{
    if (frm.imagetalk.value.length < 1) {
        alert("������ �Է��ϼ���.");
        return;
    }

    if ((frm.image1.value.length > 0) && (frm.image1.fileSize > 1000000)) {
        alert("�̹����� 1�ް��� ������ �����ϴ�.");
        return;
    }

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
        frm.submit();
    }
}

function SubmitDelete()
{
    if (confirm("������ �����Ͻðڽ��ϱ�?") == true) {
        frm.mode.value = "delete";
        frm.submit();
    }
}

function FileCheck(comp,maxfilesize,maxwidth,maxheight){
	if(comp.fileSize > maxfilesize){
		alert("���ϻ������ "+ maxfilesize + "byte�� �ѱ�� �� �����ϴ�...");
		return false;
	}

	if ((comp.src!="")&&(comp.width <1)){
		alert('�̹����� �����մϴ�.');
		return false;
	}

	return true;
}
</script>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" action="<%=uploadUrl%>/linkweb/doTopBrandTalk.asp" method=post onsubmit="return false" enctype="multipart/form-data">
    <input type=hidden name=menupos value="<%= menupos %>">
    <input type=hidden name=mode value="modify">
    <input type=hidden name=idx value="<%= idx %>">
    
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>�귣����ũ ����</b></font>
			&nbsp;
			<% if (iscurrtopbrand = false) then %>
        	<font color=red><b>���� ž�귣�尡 �ƴմϴ�.</b></font>
			<% end if %>
		</td>
	</tr>
    
    <tr align="center">
        <td valign="top" width="80" bgcolor="<%= adminColor("tabletop") %>">�����Է�</td>
        <td align="left" bgcolor="#FFFFFF">
        	<textarea class="textarea" name="imagetalk" cols="75" rows="5"><%= db2html(otopbrandtalklist.FOneItem.Fimagetalk) %></textarea>
        </td>
	</tr>
    <tr align="center">
        <td bgcolor="<%= adminColor("tabletop") %>">�̹������</td>
        <td align="left" bgcolor="#FFFFFF">
        	<input type="file" class="file" name=image1 size=40>
        	<br>
        	(1�ް� ������ �̹����� ���ε尡 �����մϴ�. ���ε����� �����ø� �����̹����� ���� �˴ϴ�.)
        </td>   
	</tr>
	</form>
	
	<form name="f" action="brandtalk_write.asp" method=get onsubmit="return false">
    <input type=hidden name=menupos value="<%= menupos %>">	
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<input type="button" class="button" value="����ϱ�" onClick="SubmitWrite();">
			<input type="button" class="button" value="�����ϱ�" onClick="SubmitDelete();">
			<input type="button" class="button" value="����ϱ�" onClick="history.back();">
        </td>
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->
<%

set otopbrandtalklist = Nothing

%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->