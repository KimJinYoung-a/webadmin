<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/company/nv/incGlobalVariable.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/company/between/betweenCls.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body>
<%
Dim cDisp, i, vDepth, vCurrpage, vPageSize, vParam, vSearch, dispCate
vCurrPage	= NullFillWith(Request("cpg"), "1")
vDepth 		= NullFillWith(Request("depth_s"), "1")
vPageSize	= NullFillWith(Request("pagesize"), 20)
vSearch		= Request("search")
dispCate	= Request("disp")

Dim itemid, itemname, sellyn, limityn, sailyn, sortDiv, sortDivOrdMall, bwdisplay
Dim schBetCateCD
itemid			= request("itemid")
itemname		= request("itemname")
sellyn			= request("sellyn")
limityn			= request("limityn") 
sailyn			= request("sailyn")
sortDiv			= request("sortDiv")
sortDivOrdMall	= request("sortDivOrdMall")
schBetCateCD	= request("schBetCateCD")
bwdisplay		= request("bwdisplay")

SET cDisp = New cDispCate
	cDisp.FCurrPage					= vCurrpage
	cDisp.FPageSize					= vPageSize
	cDisp.FRectDepth				= vDepth
	cDisp.FRectItemID 				= itemid
	cDisp.FRectItemName			 	= itemname
	cDisp.FRectSellYN				= sellyn
	cDisp.FRectLimityn				= limityn
	cDisp.FRectSailYn				= sailyn
	If (sortDiv = "on") Then
	    cDisp.FRectSortDiv			= "B"
	ElseIf (sortDivOrdMall = "on") Then
	    cDisp.FRectSortDiv			= "BM"
	End If
	cDisp.FSchBetCateCD				= schBetCateCD
	cDisp.FRectbwdisplay			= bwdisplay
	cDisp.GetRegedItemList()
%>
<script language='javascript'>
function goPage(pg){
    document.frmitem.cpg.value = pg;
    document.frmitem.submit();
}
function chgname(it){
	var popwin=window.open('/admin/etc/between/reged/pop_chgItemname.asp?itemid='+it+'','pop_chgItemname','width=500,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function checkComp(comp){
    if ((comp.name=="sortDiv")||(comp.name=="sortDivOrdMall")){
        if ((comp.name=="sortDiv")&&(comp.checked)){
            comp.form.sortDivOrdMall.checked=false;
        }

        if ((comp.name=="sortDivOrdMall")&&(comp.checked)){
            comp.form.sortDiv.checked=false;
        }
    }
}
function BetweenIsDisplay(chkYn){
	var chkSel=0, strSell;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
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

	switch(chkYn) {
		case "Y": strSell="������";break;
		case "N": strSell="���þ���";break;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� ���ÿ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditDisplay";
        document.frmSvArr.isdisplay.value = chkYn;
        document.frmSvArr.action = "/admin/etc/between/reged/reged_proc.asp"
        document.frmSvArr.submit();
    }
}
function Check_All()
{
	var chk = document.frmSvArr.cksel; 
	var cnt = 0;
	var ischecked = ""
	if(document.getElementById("chkall").checked){
		ischecked = "checked"
	}else{
		ischecked = ""
	}
	if(cnt == 0 && chk.length != 0){
		for(i = 0; i < chk.length; i++){ chk.item(i).checked = ischecked; }
		cnt++;
	}
}
</script>
<table width="700" border="0" class="a">
<tr>
	<td>&gt;&gt;��ǰ��ȸ</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<form name="frmitem" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="search" value="o">
<input type="hidden" name="cpg" value="1">
<tr>
	<td class="a">
		ī�װ� : <%= fnStandardDispCateSelectBox("1", "", "schBetCateCD", schBetCateCD, "") %>
		<br>
		��ǰ�ڵ�: <input type="text" name="itemid" value="<%= itemid %>" size="60" class="text"> &nbsp;
		�ٹ����� ��ǰ��: <input type="text" name="itemname" value="<%= itemname %>" size="50" class="text">
		<br>
		<input type="checkbox" name="sortDiv" <%= ChkIIF(sortDiv="on","checked","") %> onClick="checkComp(this)" ><b>����Ʈ��</b>
		&nbsp;
		<input type="checkbox" name="sortDivOrdMall" <%= ChkIIF(sortDivOrdMall="on","checked","") %> onClick="checkComp(this)" ><b>����Ʈ��(��Ʈ��)</b>
		&nbsp;
		�Ǹſ��� :
		<select name="sellyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
		</select>
		&nbsp;
		�������� :
		<select name="limityn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >�Ϲ�
		</select>
		&nbsp;
		���Ͽ��� :
		<select name="sailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >���ξ���
		</select>
		&nbsp;
		��Ʈ�� ���ÿ��� :
		<select name="bwdisplay" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(bwdisplay="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(bwdisplay="N","selected","") %> >���þ���
		</select>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frmitem.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="isdisplay" value="">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(vCurrPage,0) %> / <%= FormatNumber(cDisp.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(cDisp.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="30">
	<td>�̹���</td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��<br>��ǰ��</td>
	<td>��Ʈ�� ��ǰ��</td>
	<td>�ٹ�����<br>�ǸŰ�</td>
	<td>ǰ������</td>
	<td>��Ʈ��<br>���ÿ���</td>
	<td>��Ʈ�� ī�װ�</td>
	<td>3���� �Ǹŷ�</td>
</tr>
<%
If cDisp.FResultCount = 0 Then
%>
	<tr>
		<td colspan="11" height="30" bgcolor="#FFFFFF" align="center">�˻��� ��ǰ�� �����ϴ�.</td>
	</tr>
<%
Else
	For i=0 To cDisp.FResultCount-1
%>
	<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
		<td align="center"><img src="<%=cDisp.FItemList(i).FSmallImage%>"></td>
		<td align="center">
			<%=cDisp.FItemList(i).FItemID%>
			<% if cDisp.FItemList(i).FLimitYn="Y" then %><br><%= cDisp.FItemList(i).getLimitHtmlStr %></font><% end if %>
		</td>
		<td><%=cDisp.FItemList(i).FMakerID%> <%= cDisp.FItemList(i).getDeliverytypeName %> <br><%=cDisp.FItemList(i).FItemName%></td>
		<td><font Color="RED"><%=cDisp.FItemList(i).FChgItemname%></font></td>
		<td align="center">
	        <% if cDisp.FItemList(i).FSaleYn="Y" then %>
	        <strike><%= FormatNumber(cDisp.FItemList(i).FOrgPrice,0) %></strike><br>
	        <font color="#CC3333"><%= FormatNumber(cDisp.FItemList(i).FSellcash,0) %></font>
	        <% else %>
	        <%= FormatNumber(cDisp.FItemList(i).FSellcash,0) %>
	        <% end if %>
		</td>
		<td align="center">
	    <% If cDisp.FItemList(i).IsSoldOut Then %>
	        <% If cDisp.FItemList(i).FSellyn = "N" Then %>
	        <font color="red">ǰ��</font>
	        <% Else %>
	        <font color="red">�Ͻ�ǰ��</font>
	        <% End If %>
	    <% End If %>
		</td>
		<td align="center"><%= cDisp.FItemList(i).FIsdisplay %></td>
		<td>
			<span style="font-size:0.9em"><%=fnCateCodeNameSplitNotlink(cDisp.FItemList(i).FCateName,cDisp.FItemList(i).FItemID)%></span>
		</td>
		<td><%= cDisp.FItemList(i).FRctSellCNT %></td>
	</tr>
<%
	Next
%>
	<tr height="50" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if cDisp.HasPreScroll then %>
			<a href="javascript:goPage('<%= cDisp.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + cDisp.StartScrollPage to cDisp.FScrollCount + cDisp.StartScrollPage - 1 %>
    			<% if i>cDisp.FTotalpage then Exit for %>
    			<% if CStr(vCurrpage)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if cDisp.HasNextScroll then %>
    			<a href="javascript:goPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
<%
End If
%>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<% SET cDisp = nothing %>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->