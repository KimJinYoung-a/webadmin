<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/topbrand/topbrandnewscls.asp" -->
<%

dim i, page, iscurrtopbrand

page = requestCheckVar(request("page"),10)
if (page = "") then
    page = "1"
end if


'==============================================================================
dim otopbrandnewslist
set otopbrandnewslist = New CTopBrandNews

iscurrtopbrand = otopbrandnewslist.IsCurrentTopBrand(session("ssBctId"))

otopbrandnewslist.FRectMakerID = session("ssBctId")
otopbrandnewslist.FCurrPage = page
'otopbrandnewslist.FRectIsCurrentTopBrand = "Y"

otopbrandnewslist.GetTopBrandNewsList

%>
<script>
function SubmitWrite()
{
    if (frm.title.value.length < 1) {
        alert("������ �Է��ϼ���.");
        return;
    }

    if (frm.contents.value.length < 1) {
        alert("������ �Է��ϼ���.");
        return;
    }

    if (confirm("�����Ͻðڽ��ϱ�?") == true) {
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
	<form name="frm" action="brandnews_process.asp" method=post onsubmit="return false">
    <input type=hidden name=menupos value="<%= menupos %>">
    <input type=hidden name=mode value="write">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>�귣�崺�� ���</b></font>
			&nbsp;
			<% if (iscurrtopbrand = false) then %>
	        	<font color=red><b>���� ž�귣�尡 �ƴմϴ�.</b></font>
			<% end if %>
		</td>
	</tr>
	<tr align="center">
        <td width="80" bgcolor="<%= adminColor("tabletop") %>">����</td>
        <td align="left" bgcolor="#FFFFFF">
        	<input type="text" class="text" name="title" size=75>
        </td>
	</tr>
	<tr align="center">
        <td bgcolor="<%= adminColor("tabletop") %>">����</td>
        <td align="left" bgcolor="#FFFFFF">
        	<textarea class="textarea" name="contents" cols="75" rows="5"></textarea>
        </td>        
	</tr>
	</form>

	<form name="f" action="brandnews_write.asp" method=get onsubmit="return false">
    <input type=hidden name=menupos value="<%= menupos %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
          <input type="button" class="button" value="����ϱ�" onClick="SubmitWrite();">
          <input type="button" class="button" value="����ϱ�" onClick="history.back();">
	    </td>
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->
<%

set otopbrandnewslist = Nothing

%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->