<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_hot_managecls.asp" -->
<%
dim idx, poscode,reload, cdl, cdm , cds , wid , uid
Dim  i , yyyy1 , mm1
idx = request("idx")
poscode = request("poscode")
reload = request("reload")
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds") '2012 �߰� : ����ȭ

wid = session("ssBctId") '�α���ID

if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oCateContents
set oCateContents = new CCateContents
oCateContents.FRectIdx = idx
oCateContents.GetOneCateiIemContents

If cdl = "" Then
	cdl = oCateContents.FOneItem.Fcdl
End IF

If cdm = "" AND cdl = oCateContents.FOneItem.Fcdl Then
	cdm = oCateContents.FOneItem.Fcdm
End If

'2012-04-03 ����ȭ �߰�
If cds = "" AND cdl = oCateContents.FOneItem.Fcdl AND cdm = oCateContents.FOneItem.Fcdm Then
	cds = oCateContents.FOneItem.Fcds
End If

dim oposcode

'����
If yyyy1 = "" Then yyyy1 = Year(Now())
If mm1 = "" Then mm1 = Month(Now())

%>

<script language='javascript'>
function SaveCateContents(frm){
    if (frm.cdl.value == ""){
       alert('��ī�װ��� �Է� �ϼ���.');
        frm.cdl.focus();
        return;
    }

    if (frm.cdl.value == "110" && frm.cdm.value == ""){
       alert('��ī�װ��� �Է� �ϼ���.');
        frm.cdm.focus();
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

// ī�װ� ����� ���
function changecontent(){
	<% If oCateContents.FOneItem.Fidx <> "" Then %>
		alert("ī�װ��� ������ �� �� name�� Map_hot_ �� �ڵ尪(����ī�װ���)�� ����� �����ؾ� �մϴ�. ");
		document.getElementById("categorylist").style.display = "block";
	<% Else %>
		location.href = "?cdl=" + frmcontents.cdl.value + "<%=CHKIIF(cdl<>"" and cdl = 110 ,"&cdm="&chr(34)&" + frmcontents.cdm.value + "&chr(34)&"","")%>&idx=<%=idx%>";
	<% End If %>
}
</script>
<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="volume_Proc.asp" onsubmit="return false;">
<input type="hidden" name="wid" value="<%=wid%>">
<tr bgcolor="#FFFFFF">
    <td width="20%" bgcolor="#DDDDFF">Idx</td>
    <td >
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        <%= oCateContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oCateContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>

<tr bgcolor="#FFFFFF">
    <td  bgcolor="#DDDDFF">ī�װ�</td>
    <td>
    	<font color="red">�� ��ī�װ��� �Է� ���ּ��� (����ä���� ��ī�װ�)</font><br>
        <%
        	if oCateContents.FOneItem.Fidx<>"" then
        		call DrawSelectBoxCategoryLarge("cdl", cdl)
        		Response.Write "&nbsp;"
        		if cdl = "110" And cdl <> "" Then '����ä���ϰ��
        			call DrawSelectBoxCategoryMid("cdm",cdl, cdm)
					Response.Write "&nbsp;"
'					If cdm <> "" Then
'						call DrawSelectBoxCategorySmall("cds",cdl, cdm, cds )
'					End If 
        		end if
        	else
    			call DrawSelectBoxCategoryLarge("cdl", cdl)
    			Response.Write "&nbsp;"
    			if cdl = "110" And cdl <> "" Then '����ä���ϰ��
    				call DrawSelectBoxCategoryMid("cdm",cdl, cdm)
					Response.Write "&nbsp;"
'					If cdm <> "" Then
'						call DrawSelectBoxCategorySmall("cds",cdl, cdm, cds )
'					End If 
    			end if
        	end if
        %>
        <br>
    </td>
</tr>

<tr bgcolor="#FFFFFF">
    <td width="20%" bgcolor="#DDDDFF">��¥</td>
    <td >
		<select class="select" name="yyyy1">
		<%
			for i=2002 to Year(now)
				if (CStr(i)=CStr(yyyy1)) Then
		%>
				<option value="<%=CStr(i)%>" selected><%=CStr(i)%></option>
		<% Else %>
				<option value="<%=CStr(i)%>" ><%=CStr(i)%></option>
		<%
				end if
			next
		%>
		</select>��
		<select class="select" name="mm1">
		<%
			for i=1 to 12
				if (Format00(2,i)=Format00(2,mm1)) Then
		%>
				<option value="<%=Format00(2,i)%>" selected><%=Format00(2,i)%></option>
		<% Else %>
				<option value="<%=Format00(2,i)%>" ><%=Format00(2,i)%></option>
		<%
				end if
			next
		%>
		</select>��
    </td>
</tr>


<tr bgcolor="#FFFFFF">
    <td  bgcolor="#DDDDFF">��ǥ�ŷ���</td>
    <td>
		<input type="text" name="volume" value="">��
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td  bgcolor="#DDDDFF">��ǥ���;�</td>
    <td>
		<input type="text" name="revenus" value="">��
	</td>
</tr>

<tr bgcolor="#FFFFFF">
    <td  bgcolor="#DDDDFF">�����</td>
    <td>
        <%= oCateContents.FOneItem.Fregdate %> (<%= oCateContents.FOneItem.Freguserid %>)
    </td>
</tr>

<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center" style="padding:5 0 5 0">
	    <table cellpadding="0" cellspacing="0" border="0">
			<tr><td style="padding-bottom:5px;"></td></tr>
	    </table>
		<input type="button" value=" �� �� " onClick="SaveCateContents(frmcontents);">
    </td>
</tr>
</form>
</table>

<script language="JavaScript">
<!--
var speed = 100 //�����̴� �ӵ� - 1000�� 1��

function doBlink(){
var blink = document.all.tags("blink")
for (var i=0; i < blink.length; i++)
blink[i].style.visibility = blink[i].style.visibility == "" ? "hidden" : ""
} 

function startBlink() { 
setInterval("doBlink()",speed)
} 
window.onload = startBlink; 
//-->
</script>

<%
set oposcode = Nothing
set oCateContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
