<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_hot_managecls.asp" -->
<%
dim idx, poscode,reload, cdl, cdm , cds
idx = request("idx")
poscode = request("poscode")
reload = request("reload")
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds") '2012 �߰� : ����ȭ

if idx="" then idx=0

if reload="on" then
    response.write "<script>opener.location.reload(); window.close();</script>"
    dbget.close()	:	response.End    
end if

dim oCateContents
set oCateContents = new CCateContents
oCateContents.FRectIdx = idx
oCateContents.GetOneCateContents

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

dim oposcode, defaultMapStr

    defaultMapStr = "<map name='Map_hot_" & cdl & cdm & "'>" + VbCrlf
    defaultMapStr = defaultMapStr + VbCrlf
    defaultMapStr = defaultMapStr + "</map>"

%>

<script language='javascript'>
function SaveCateContents(frm){
    if (frm.cdl.value == ""){
       alert('��ī�װ����� �Է� �ϼ���.');
        frm.cdl.focus();
        return;
    }
    
    if (frm.cdm.value == ""){
       alert('��ī�װ����� �Է� �ϼ���.');
        frm.cdm.focus();
        return;
    }
    
//    if (frm.linktype.value == ""){
//       alert('��ũ ������ �Է� �ϼ���.');
//        frm.linktype.focus();
//        return;
//    }
//    
//    if (frm.linktype.value == "F" || frm.linktype.value == "X"){
//        alert('��ũ ������\n\n��ũ (a href)\n�� (#Map)\n\n�� �����ϼ���.');
//        frm.linktype.focus();
//        return;
//    }
//    
//    if (frm.linkurl.value.length<1){
//        alert('��ũ ���� �Է� �ϼ���.');
//        frm.linkurl.focus();
//        return;
//    }
    
    if (frm.startdate.value.length!=10){
        alert('�������� �Է�  �ϼ���.');
        frm.startdate.focus();
        return;
    }
    
    if (frm.enddate.value.length!=10){
        alert('�������� �Է�  �ϼ���.');
        frm.enddate.focus();
        return;
    }
    
    var vstartdate = new Date(frm.startdate.value.substr(0,4), frm.startdate.value.substr(5,2), frm.startdate.value.substr(8,2));
    var venddate = new Date(frm.enddate.value.substr(0,4), frm.enddate.value.substr(5,2), frm.enddate.value.substr(8,2));
    
    if (vstartdate>venddate){
        alert('�������� �����Ϻ��� Ŭ �� �����ϴ�.');
        frm.enddate.focus();
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

// ī�װ��� ����� ����
function changecontent(){
	<% If oCateContents.FOneItem.Fidx <> "" Then %>
		alert("ī�װ����� ������ �� �� name�� Map_hot_ �� �ڵ尪(����ī�װ�����)�� ����� �����ؾ� �մϴ�. ");
		document.getElementById("categorylist").style.display = "block";
	<% Else %>
		location.href = "?cdl=" + frmcontents.cdl.value + "<%=CHKIIF(cdl<>"","&cdm="&chr(34)&" + frmcontents.cdm.value + "&chr(34)&"","")%><%=CHKIIF(cdm<>"","&cds="&chr(34)&" + frmcontents.cds.value + "&chr(34)&"","")%>&idx=<%=idx%>";
	<% End If %>
}
</script>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/Category/doCateHotReg.asp" onsubmit="return false;" enctype="multipart/form-data">
<tr bgcolor="#FFFFFF">
    <td width="20%" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        <%= oCateContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oCateContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">ī�װ���</td>
    <td>
        <%
        	if oCateContents.FOneItem.Fidx<>"" then
        		call DrawSelectBoxCategoryLarge("cdl", cdl)
        		Response.Write "&nbsp;"
        		if cdl <> "" then
        			call DrawSelectBoxCategoryMid("cdm",cdl, cdm)
					Response.Write "&nbsp;"
					If cdm <> "" Then
						call DrawSelectBoxCategorySmall("cds",cdl, cdm, cds )
					End If 
        		end if
        	else
    			call DrawSelectBoxCategoryLarge("cdl", cdl)
    			Response.Write "&nbsp;"
    			if cdl <> "" then
    				call DrawSelectBoxCategoryMid("cdm",cdl, cdm)
					Response.Write "&nbsp;"
					If cdm <> "" Then
						call DrawSelectBoxCategorySmall("cds",cdl, cdm, cds )
					End If 
    			end if
        	end if
        %>
        <br>
        <div id="categorylist" style="display:none;">
        <%
			   dim tmp_str,query1, tt
			   tt = 1
			
			   query1 = " select code_mid, code_nm from [db_item].[dbo].tbl_Cate_mid"
			   query1 = query1 & " where display_yn = 'Y'"
			   query1 = query1 & " and code_large = '" & cdl & "'"
			   query1 = query1 & " and code_mid<>0"
			   query1 = query1 & " order by code_mid Asc"
			
			   rsget.Open query1,dbget,1
			
			   if  not rsget.EOF  then
			       rsget.Movefirst
			
			       do until rsget.EOF
			           response.write("["&rsget("code_mid")&"]"& db2html(rsget("code_nm")) &",")
			           if tt = 5 then response.write "<br>" end if
			           rsget.MoveNext
			           tt = tt + 1
			       loop
			   end if
			   rsget.close
        %>
        </div>
    </td>
</tr>
<!-- <tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��ũ����</td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
        	<%= DrawLinktypeCombo("",oCateContents.FOneItem.Flinktype,"") %>
        <% else %>
            <%= DrawLinktypeCombo("","","") %>
        <% end if %>
        &nbsp;&nbsp;��ũ������ ��ũ (a href), �� (#Map) �� �����ϼ���.
    </td>
</tr> -->

<!-- <tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�̹���</td>
  <td><input type="file" name="file1" value="" size="32" maxlength="32" class="file">
  <% if oCateContents.FOneItem.Fidx<>"" then %>
  <br>
  <img src="<%= oCateContents.FOneItem.GetImageUrl %>" >
  <br> <%= oCateContents.FOneItem.GetImageUrl %>
  <% end if %>
  </td>
</tr> -->
<!-- <tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��ũ��<br><br><b><font color="red">URL�� ���<br>�� ���� �� �Է�</font></b></td>
    <td>
        <% if oCateContents.FOneItem.Fidx<>"" then %>
            <textarea name="linkurl" cols="60" rows="6"><%= oCateContents.FOneItem.Flinkurl %></textarea>
            <br><b>(�̹������� ��� name�� ���� ���� ����)</b>
        <% else %>
                <textarea name="linkurl" cols="60" rows="6"><%= defaultMapStr %></textarea>
                <br><b>(�̹������� ��� name�� ���� ���� ����)</b>
        <% end if %>
        <br><b>���ΰ�� map name='Map_hot_ �� �ݵ�� ��ī���ڵ�� ��ī���ڵ尡 ������!<br>��� �ʼҽ��� ĭ�������� �� �ٿ�����!<br>��ũ�� �ڹٽ�ũ��Ʈ �ڵ� ������ �ȵ�!</b>
    </td>
</tr> -->
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
    <td>
        <input type="text" name="startdate" value="<%= oCateContents.FOneItem.Fstartdate %>" maxlength="10" size="10"> (<a href="#" onClick="frmcontents.startdate.value='<%= Left(CStr(now()),10) %>'"><%= Left(CStr(now()),10) %></a>)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
    <td>
        <input type="text" name="enddate" value="<%= Left(oCateContents.FOneItem.Fenddate,10) %>" maxlength="10" size="10">
        <input type="text" name="dummy1" value="23:59:59" size="8" readonly style="background:'#EEEEEE'">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�����</td>
    <td>
        <%= oCateContents.FOneItem.Fregdate %> (<%= oCateContents.FOneItem.Freguserid %>)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">��뿩��</td>
    <td>
        <% if oCateContents.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">�����
        <input type="radio" name="isusing" value="N" checked >������
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >�����
        <input type="radio" name="isusing" value="N">������
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center" style="padding:5 0 5 0">
    <table cellpadding="0" cellspacing="0" border="0">
	<tr><td style="padding-bottom:5px;">�� <b><font color="red">�������䰡 ���� ī�װ���(����,����Ʈ)���� �ǽð� �����ϸ� ������ �׽��ϴ�!!</font></b></td></tr>
    <tr><td style="padding-bottom:5px;">�� <b><font color="blue">�������� ���ñ������� �Ϸ������� ���ּ���.!! �������� ���÷� �ϸ� ������ �ȵ�.</font></b></td></tr>
    <tr><td style="padding-bottom:5px;">�� <b><font color="red">�� <font color="blue">���ó�¥�� ���� 6�� ��������</font> �ø��°��� ����˴ϴ�.</font></b></td></tr>
    <tr><td style="padding-bottom:15px;">�� <b><font color="red">��ũ ������ Ȯ���� ���ּ���!!</font> �׳� <font color="blue">��ũ�϶��� ��ũ</font>��, <font color="blue">�� �ϰ�� ��</font>���� ����!!!</b></td></tr>
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