<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Dim midx, page, i, mallid, setMargin, itemid
page		= request("page")
midx 		= request("midx")
mallid		= request("mallid")
setMargin	= request("setMargin")
itemid  	= request("itemid")

If page = "" Then page = 1

If NOT isNumeric(midx) Then
	Response.Write "<script language=javascript>alert('�߸��� �����Դϴ�.');window.close();</script>"
	dbget.close()	:	response.End
End If

If mallid = "" Then
	Response.Write "<script language=javascript>alert('�߸��� �����Դϴ�.');window.close();</script>"
	dbget.close()	:	response.End
End If

If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

Dim oSsg
Set oSsg = new Cssg
	oSsg.FCurrPage			= page
	oSsg.FPageSize			= 50
	oSsg.FRectMallid		= mallid
	oSsg.FRectMasterIdx		= midx
	oSsg.FRectsetMargin		= setMargin
	oSsg.FRectItemID		= itemid
	oSsg.getssgMarginItemDetailList
%>
<link rel="stylesheet" href="/bct.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function popCateSelect(){
	$.ajax({
		url: "/admin/etc/ssg/act_CategorySelect.asp",

		cache: false,
		success: function(message) {
			$("#lyrCateAdd").empty().append(message).fadeIn();
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function jsAddItemID() {
	var frm = document.frm;

	if (frm.itemid.value == '') {
		alert('��ǰ�ڵ带 �Է��ϼ���.');
		return;
	}

	if (confirm('�����Ͻðڽ��ϱ�?')) {
		frm.delIdx.value = '';
		frm.submit();
	}
}

function delItem(v)
{
	$("#delIdx").val(v);
	document.frm.submit();
}


function selectDeleteProcess() {
	var chkSel=0;
	try {
		if(frmlist.cksel.length>1) {
			for(var i=0;i<frmlist.cksel.length;i++) {
				if(frmlist.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmlist.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
		document.frmlist.mode.value = "selDel";
		document.frmlist.action = "/admin/etc/ssg/procSsgMargin.asp";
		document.frmlist.submit();
    }
}

function goPage(pg){
    //frm.page.value = pg;
    //frm.submit();
	location.href = '?page='+pg+'&midx=<%= midx %>&mallid=<%=mallid%>';
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="get" action="">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		���� : <input type="text" name="setMargin" value="<%= setMargin %>" class="text" size="5" maxlength="5">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
	</td>
</tr>
</form>
</table>
<br /><br />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="3" bgcolor="<%= adminColor("tabletop") %>">�Ⱓ�� ���� ����Ʈ</td>
</tr>
</table>

<br />

<form name="frm" action="procSsgMargin.asp" methd="post" style="margin:0px;">
<input type="hidden" name="mode" value="itemDetail">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<input type="hidden" id="delIdx" name="delIdx" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">
				��ǰID :
				<textarea class="textarea" name="itemid" cols="16" rows="2"></textarea>
				<input type="button" value="�� ��" onClick="jsAddItemID()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<!-- ����Ʈ ���� -->
<form name="frmlist" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="midx" value="<%= midx %>">
<input type="hidden" name="page" value="<%= page %>">
<input type="hidden" name="mallid" value="<%= mallid %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		�˻���� : <b><%= FormatNumber(oSsg.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSsg.FTotalPage,0) %></b>
	</td>
	<td align="center"><input class="button" type="button" id="btnCommcd" value="���û���" onClick="selectDeleteProcess();" ></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="2%"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmlist.cksel);"></td>
    <td width="100">IDX</td>
	<td>��ǰ�ڵ�</td>
	<td>�������븶��</td>
	<td width="100">����</td>
</tr>
<% For i=0 to oSsg.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oSsg.FItemList(i).Fidx %>"></td>
	<td><%= oSsg.FItemList(i).Fidx %></td>
	<td><%= oSsg.FItemList(i).Fitemid %></td>
	<td><%= oSsg.FItemList(i).FSetMargin %>%</td>
	<td><input type="button" class="button" value="����" onclick="delItem(<%= oSsg.FItemList(i).FIdx %>);"></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oSsg.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSsg.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSsg.StartScrollPage to oSsg.FScrollCount + oSsg.StartScrollPage - 1 %>
    		<% if i>oSsg.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSsg.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<% Set oSsg = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
