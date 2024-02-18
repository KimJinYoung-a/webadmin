<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/between/betweenItemcls.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim cDisp, i, vDepth, vCurrpage, vPageSize, vParam, vSearch, vNotCateReg, dispCate, onlyValidMargin
vCurrPage	= NullFillWith(Request("cpg"), "1")
vDepth 		= NullFillWith(Request("depth_s"), "1")
vPageSize	= NullFillWith(Request("pagesize"), 20)
vSearch		= Request("search")
vNotCateReg	= Request("notcatereg")
dispCate	= Request("disp")

Dim makerid, itemid, itemname, sellyn, limityn, sailyn, sortDiv, sortDivOrdMall, bwdisplay
Dim schBetCateCD
makerid			= request("makerid")
itemid			= request("itemid")
itemname		= request("itemname")
sellyn			= request("sellyn")
usingyn			= request("usingyn")
danjongyn		= request("danjongyn")
limityn			= request("limityn")
sailyn			= request("sailyn")
sortDiv			= request("sortDiv")
sortDivOrdMall	= request("sortDivOrdMall")
schBetCateCD	= request("schBetCateCD")
onlyValidMargin	= request("onlyValidMargin")
bwdisplay		= request("bwdisplay")

Dim cPjtGroup, pjt_code
pjt_code = request("pjt_code")

SET cDisp = New cDispCate
	cDisp.FCurrPage					= vCurrpage
	cDisp.FPageSize					= vPageSize
	cDisp.FRectDepth				= vDepth
	cDisp.FRectMakerId 				= makerid
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
	cDisp.FRectNotCateReg			= vNotCateReg
	cDisp.FSchBetCateCD				= schBetCateCD
	cDisp.FRectonlyValidMargin		= onlyValidMargin
	cDisp.FRectbwdisplay			= bwdisplay
	cDisp.GetRegedItemList()
%>
<script language='javascript'>
function goPage(pg){
    document.frmitem.cpg.value = pg;
    document.frmitem.submit();
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
function SelectItems(sType){
	var itemcount = 0;
	var frm;
	frm = document.frmSvArr;
	frm.sType.value = sType;

	if (sType == "sel"){
		if(typeof(frm.cksel) !="undefined"){
			if(!frm.cksel.length){
				if(!frm.cksel.checked){
					alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
					return;
				}
				frm.itemidarr.value = frm.cksel.value;
			}else{
				for(i=0;i<frm.cksel.length;i++){
					if(frm.cksel[i].checked) {
						if (frm.itemidarr.value==""){
							frm.itemidarr.value =  frm.cksel[i].value;
						}else{
							frm.itemidarr.value = frm.itemidarr.value + "," +frm.cksel[i].value;
						}
					}
				}
				if (frm.itemidarr.value == ""){
					alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
					return;
				}
			}
		}else{
			alert("�߰��� ��ǰ�� �����ϴ�.");
			return;
		}
	}
	frm.target = "FrameCKP";
	frm.action = "/admin/etc/between/project/projectitem_process.asp";
	frm.submit();
	frm.itemidarr.value = "";
	opener.history.go(0);
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<form name="frmitem" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="search" value="o">
<input type="hidden" name="cpg" value="1">
<input type="hidden" name="pjt_code" value="<%= pjt_code %>">
<tr>
	<td class="a">
		�� �� �� : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		�ٹ����� ��ǰ��: <input type="text" name="itemname" value="<%= itemname %>" size="50" class="text">
		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >���� <%= CMAXMARGIN %>%�̻� ��ǰ�� ����
		<br>
		ī�װ� : <%= fnStandardDispCateSelectBox("1", "", "schBetCateCD", schBetCateCD, "") %>
		<br>
		��ǰ��ȣ: <input type="text" name="itemid" value="<%= itemid %>" size="60" class="text"> &nbsp;
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
<%
SET cPjtGroup = new cProject
	cPjtGroup.FRectPjt_code = pjt_code
	cPjtGroup.getProjectItemGroup()
%>
<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="isdisplay" value="">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="sType" >
<input type="hidden" name="itemidarr" >
<input type="hidden" name="mode" value="I">
<tr>
	<td  valign="bottom">
		<select name="selGroup">
			<option value="0"> �׷� ������ </option>
	<%
		If cPjtGroup.FResultCount > 0 Then
			For i = 0 to cPjtGroup.FResultCount - 1
	%>
			<option value="<%=cPjtGroup.FItemList(i).FPjtgroup_code%>" ><%IF cPjtGroup.FItemList(i).FPjtgroup_pcode <> 0 THEN%>��&nbsp;<%END IF%><%=cPjtGroup.FItemList(i).FPjtgroup_code%>(<%=cPjtGroup.FItemList(i).FPjtgroup_desc%>)</option>
	<%
			Next
		END IF
	%>
	 	</select>
		<input type="button" value="���û�ǰ �߰�" onClick="SelectItems('sel')" class="button">
	</td>
</tr>
</table>
<%
SET cPjtGroup = nothing
%>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(vCurrPage,0) %> / <%= FormatNumber(cDisp.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(cDisp.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="30">
	<td><input type="checkbox" name="chkall" id="chkall" value="" onClick="Check_All()"></td>
	<td>�̹���</td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��<br>��ǰ��</td>
	<td>��Ʈ��<br>���� ��ǰ��</td>
	<td>�ٹ�����<br>�ǸŰ�</td>
	<td>��Ʈ��<br>���ÿ���</td>
	<td>�ٹ�����<br>����</td>
	<td>�ٹ�����<br>����ī�װ�</td>
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
		<td align="center"><input type="checkbox" name="cksel" value="<%=cDisp.FItemList(i).FItemID%>"></td>
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
		<td align="center"><%= cDisp.FItemList(i).FIsdisplay %></td>
		<td align="center">
	        <% if cDisp.FItemList(i).Fsellcash<>0 then %>
				<%= CLng(10000-cDisp.FItemList(i).Fbuycash/cDisp.FItemList(i).Fsellcash*100*100)/100 %> %
	        <% end if %>
		</td>
		<td>
			<span style="font-size:0.9em"><%=fnCateCodeNameSplit2(cDisp.FItemList(i).FCateName2,cDisp.FItemList(i).FItemID)%></span>
		</td>
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
<% SET cDisp = nothing %>
<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->