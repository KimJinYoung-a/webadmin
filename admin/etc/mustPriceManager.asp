<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/mustPriceCls.asp"-->
<%
Dim makerid, itemid, mallgubun, isGetDate
Dim page, i, mwdiv
Dim oMustPrice
page                = request("page")
makerid				= requestCheckVar(request("makerid"), 32)
itemid  			= request("itemid")
mallgubun           = requestCheckVar(request("mallgubun"), 32)
isGetDate           = requestCheckVar(request("isGetDate"), 1)
mwdiv				= request("mwdiv")

If page = "" Then page = 1
'�ٹ����� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
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

SET oMustPrice = new CMustPrice
	oMustPrice.FCurrPage					= page
	oMustPrice.FPageSize					= 50
    oMustPrice.FRectMakerid					= makerid
	oMustPrice.FRectItemID					= itemid
    oMustPrice.FRectMallgubun				= mallgubun
    oMustPrice.FRectIsGetDate		    	= isGetDate
	oMustPrice.FRectMwdiv		    		= mwdiv
    oMustPrice.getMustPirceItemList
%>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function popMustPrice(){
	var popMustPrice = window.open("/admin/etc/popMustPrice.asp","popMustPrice","width=700,height=400,scrollbars=yes,resizable=yes");
	popMustPrice.focus();
}
function fnModifyMustPrice(iidx, mallid){
	var popMustPrice = window.open("/admin/etc/popMustPrice.asp?idx="+iidx+"&isModify=Y&mallid="+mallid,"popMustPrice","width=700,height=400,scrollbars=yes,resizable=yes");
	popMustPrice.focus();
}
function popUploadExcel(){
	var popwin;
	popwin = window.open("/admin/etc/mustprice/popUploadMustPrice.asp", "popup_item", "width=500,height=230,scrollbars=yes,resizable=yes");
	popwin.focus();
}
function fnDelItems(){
	var chkSel=0;
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
	if (confirm('�����Ͻ� ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.mode.value = "D";
		document.frmSvArr.action = "/admin/etc/mustPrice_process.asp"
		document.frmSvArr.submit();
    }
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<br /><br />
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
        �� ���� :
        <select name="mallgubun" class="select">
            <option value="">-Choice-</option>
            <option value="ssg" <%= CHKiif(mallgubun="ssg","selected","") %> >SSG</option>
            <option value="coupang" <%= CHKiif(mallgubun="coupang","selected","") %> >����</option>
            <option value="halfclub" <%= CHKiif(mallgubun="halfclub","selected","") %> >����Ŭ��</option>
			<option value="hmall1010" <%= CHKiif(mallgubun="hmall1010","selected","") %> >HMall</option>
            <option value="auction1010" <%= CHKiif(mallgubun="auction1010","selected","") %> >����</option>
            <option value="ezwel" <%= CHKiif(mallgubun="ezwel","selected","") %> >���������</option>
            <option value="gmarket1010" <%= CHKiif(mallgubun="gmarket1010","selected","") %> >G����</option>
            <option value="gsshop" <%= CHKiif(mallgubun="gsshop","selected","") %> >GSShop</option>
            <option value="interpark" <%= CHKiif(mallgubun="interpark","selected","") %> >������ũ</option>
            <option value="nvstorefarm" <%= CHKiif(mallgubun="nvstorefarm","selected","") %> >�������</option>
			<option value="Mylittlewhoopee" <%= Chkiif(mallgubun = "Mylittlewhoopee", "selected", "") %>>������� Ĺ�ص�</option>
			<option value="nvstoregift" <%= CHKiif(mallgubun="nvstoregift","selected","") %> >������� �����ϱ�</option>
            <option value="WMP" <%= CHKiif(mallgubun="WMP","selected","") %> >������</option>
			<option value="11st1010" <%= CHKiif(mallgubun="11st1010","selected","") %> >11����</option>
            <option value="lotteCom" <%= CHKiif(mallgubun="lotteCom","selected","") %> >�Ե�����</option>
            <option value="lotteimall" <%= CHKiif(mallgubun="lotteimall","selected","") %> >�Ե����̸�</option>
			<option value="lotteon" <%= CHKiif(mallgubun="lotteon","selected","") %> >�Ե�On</option>
			<option value="skstoa" <%= CHKiif(mallgubun="skstoa","selected","") %> >SKSTOA</option>
			<option value="shintvshopping" <%= CHKiif(mallgubun="shintvshopping","selected","") %> >�ż���TV����</option>
			<option value="wetoo1300k" <%= CHKiif(mallgubun="wetoo1300k","selected","") %> >1300k</option>
            <option value="cjmall" <%= CHKiif(mallgubun="cjmall","selected","") %> >CJMall</option>
			<option value="lfmall" <%= Chkiif(mallgubun = "lfmall", "selected", "") %>>LFmall</option>
			<option value="sabangnet" <%= Chkiif(mallgubun = "sabangnet", "selected", "") %>>����</option>
			<option value="kakaogift" <%= Chkiif(mallgubun = "kakaogift", "selected", "") %>>īī������Ʈ</option>
			<option value="kakaostore" <%= Chkiif(mallgubun = "kakaostore", "selected", "") %>>īī���彺���</option>
			<option value="boribori1010" <%= Chkiif(mallgubun = "boribori1010", "selected", "") %>>��������</option>
			<option value="wconcept1010" <%= Chkiif(mallgubun = "wconcept1010", "selected", "") %>>W����</option>
			<option value="benepia1010" <%= Chkiif(mallgubun = "benepia1010", "selected", "") %>>�����Ǿ�</option>
        </select>
        &nbsp;
        Ư�����࿩��(���糯¥����) :
        <select name="isGetDate" class="select">
            <option value="" >-Choice-</option>
            <option value="Y" <%= CHKiif(isGetDate="Y","selected","") %> >������</option>
        </select>
		�ŷ����� <% drawSelectBoxMWU "mwdiv", mwdiv %>
    </td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<br />
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="mallgubun" value="<%= mallgubun %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		�˻���� : <b><%= FormatNumber(oMustPrice.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMustPrice.FTotalPage,0) %></b>
	</td>
	<td align="right" colspan="2">
        <input type="button" class="button" value="����" onclick="popMustPrice();" />
		&nbsp;
        <input type="button" class="button" value="�������" onclick="popUploadExcel();" />
    <% If mallgubun <> "" Then %>
        &nbsp;
        <input type="button" class="button" value="����" onclick="fnDelItems();" />
    <% End If %>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="70">������</td>
    <td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
    <td width="200">Ư���Ⱓ</td>
    <td width="70">Ư��</td>
	<td width="70">Ư����<br>����</td>
	<td width="70">Ư����<br>���԰�</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ŷ�����</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="80">������ID</td>
</tr>
<% For i = 0 To oMustPrice.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oMustPrice.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oMustPrice.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= oMustPrice.FItemList(i).FMallgubun %></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oMustPrice.FItemList(i).FItemID%>" target="_blank"><%= oMustPrice.FItemList(i).FItemID %></a>
	</td>
	<td align="left" style="cursor:pointer;" onclick="fnModifyMustPrice('<%= oMustPrice.FItemList(i).FIdx %>', '<%= oMustPrice.FItemList(i).FMallgubun %>');">
        <%= oMustPrice.FItemList(i).FMakerid %><%= oMustPrice.FItemList(i).getDeliverytypeName %><br><%= oMustPrice.FItemList(i).FItemName %>
    </td>
	<td>
		<%= FormatDate(oMustPrice.FItemList(i).FStartDate,"0000-00-00 00:00:00") %> <br />~ <%= FormatDate(oMustPrice.FItemList(i).FEndDate,"0000-00-00 00:00:00") %>
	</td>
	<td align="right">
		<%= FormatNumber(oMustPrice.FItemList(i).FMustPrice,0) %>
	</td>
	<td align="right">
	<%
		If oMustPrice.FItemList(i).FMustMargin = 0 Then
			response.write "�����ȵ�"
		Else
			response.write oMustPrice.FItemList(i).FMustMargin & "%"
		End If
	%>
	</td>
	<td align="right">
	<%
		If oMustPrice.FItemList(i).FMustMargin = 0 Then
			response.write "�����ȵ�"
		Else
			response.write FormatNumber(oMustPrice.FItemList(i).FMustBuyPrice,0)
		End If
	%>
	</td>
	<td align="right">
	<% If oMustPrice.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oMustPrice.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oMustPrice.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oMustPrice.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oMustPrice.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oMustPrice.FItemList(i).Fbuycash/oMustPrice.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oMustPrice.FItemList(i).IsSoldOut Then
			If oMustPrice.FItemList(i).FSellyn = "N" Then
	%>
			<font color="red">ǰ��</font>
	<%
			Else
	%>
			<font color="red">�Ͻ�<br>ǰ��</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		Select Case oMustPrice.FItemList(i).FMWDiv
			Case "M"	response.write "����"
			Case "W"	response.write "��Ź"
			Case "U"	response.write "��ü"
		End Select
	%>
	</td>
	<td align="center">
	<%
		If oMustPrice.FItemList(i).FItemdiv = "06" OR oMustPrice.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center"><%= Chkiif(oMustPrice.FItemList(i).Freguserid <> "", oMustPrice.FItemList(i).Freguserid, oMustPrice.FItemList(i).FLastUpdateUserId ) %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18" align="center">
	<% If oMustPrice.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oMustPrice.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oMustPrice.StartScrollPage To oMustPrice.FScrollCount + oMustPrice.StartScrollPage - 1 %>
		<% If i>oMustPrice.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oMustPrice.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oMustPrice = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->