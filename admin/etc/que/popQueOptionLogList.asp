<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim mallid, oOutmall, page, i
Dim itemid, apiAction, resultCode, lastUserid
mallid		= request("mallid")
itemid		= request("itemid")
apiAction	= request("apiAction")
resultCode	= request("resultCode")
page 		= request("page")
lastUserid	= request("lastUserid")

If page = "" Then page = 1
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

Set oOutmall = new COutmall
	If (session("ssBctID")="kjy8517") Then
		oOutmall.FPageSize 			= 50
	Else
		oOutmall.FPageSize 			= 20
	End If
	oOutmall.FCurrPage			= page
	oOutmall.FRectMallid 		= mallid
	oOutmall.FRectItemid 		= itemid
	oOutmall.FRectApiAction 	= apiAction
	oOutmall.FRectResultCode 	= resultCode
	oOutmall.FRectLastUserid 	= lastUserid
	oOutmall.getQueOptionLogList
%>
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
// ���õ� ��ǰ �Ǹſ��� ����
function etcmallSellYnProcess(chkYn, imallid) {
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
		case "Y": strSell="�Ǹ���";break;
		case "N": strSell="ǰ��";break;
	}

	if (imallid == 'gsshop'){
	    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
	        if (chkYn=="X"){
	            if (!confirm(strSell + '�� �����ϸ� GSShop���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
	        }
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
	        document.frmSvArr.submit();
	    }
	}else if (imallid == 'lotteimall'){
		if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		    if (chkYn=="X"){
		        if (!confirm(strSell + '�� �����ϸ� �Ե�iMall���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
		    }
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EditSellYn";
			document.frmSvArr.chgSellYn.value = chkYn;
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
		}	
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td>������ : <%= mallid %></td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mallid" value="<%=mallid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		API�׼� : 
		<select name="apiAction" class="select">
			<option value="">��ü</option>
			<option value="REG"  <%= Chkiif(apiAction = "REG", "selected", "")%> >��ǰ���</option>
			<option value="EDIT"	 <%= Chkiif(apiAction = "EDIT", "selected", "")%> >��ǰ����</option>
			<option value="SOLDOUT"  <%= Chkiif(apiAction = "SOLDOUT", "selected", "")%> >ǰ��ó��</option>
			<option value="PRICE"	 <%= Chkiif(apiAction = "PRICE", "selected", "")%> >���ݼ���</option>
			<% If (mallid <> "cjmall" and mallid <> "11stmy") Then %>
			<option value="ITEMNAME" <%= Chkiif(apiAction = "ITEMNAME", "selected", "")%> >��ǰ�����</option>
			<% End If %>
			<% If mallid = "gsshop" Then %>
			<option value="IMAGE"	 <%= Chkiif(apiAction = "IMAGE", "selected", "")%> >�̹�������</option>
			<option value="CONTENT"  <%= Chkiif(apiAction = "CONTENT", "selected", "")%> >��ǰ�������</option>
			<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >���ΰ�ü���</option>
			<% ElseIf mallid = "11stmy" Then %>
			<option value="VIEWOPT"  <%= Chkiif(apiAction = "VIEWOPT", "selected", "")%> >�ɼ���ȸ</option>
			<% ElseIf mallid = "cjmall" Then %>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >�űԻ�ǰ��ȸ</option>
			<% ElseIf (mallid = "lotteimall") OR (mallid = "lotteCom") Then %>
			<option value="CHKSTOCK"  <%= Chkiif(apiAction = "CHKSTOCK", "selected", "")%> >�����ȸ</option>
			<option value="CHKSTAT"  <%= Chkiif(apiAction = "CHKSTAT", "selected", "")%> >�űԻ�ǰ��ȸ</option>
				<% If mallid = "lotteimall" Then %>
					<option value="DISPVIEW"  <%= Chkiif(apiAction = "DISPVIEW", "selected", "")%> >���û�ǰ��ȸ</option>
				<% Else %>
					<option value="INFODIV"	 <%= Chkiif(apiAction = "INFODIV", "selected", "")%> >���ΰ�ü���</option>
				<% End If %>
			<% End If %>
		</select>
		&nbsp;
		�������� : 
		<select name="resultCode" class="select">
			<option value="">��ü</option>
			<option value="OK"  	<%= Chkiif(resultCode = "OK", "selected", "")%> >����</option>
			<option value="ERR"		<%= Chkiif(resultCode = "ERR", "selected", "")%> >����</option>
			<option value="QNull"	<%= Chkiif(resultCode = "QNull", "selected", "")%> >����</option>
		</select>
		&nbsp;
		����ID : 
		<select name="lastUserid" class="select">
			<option value="">��ü</option>
			<option value="system"	<%= Chkiif(lastUserid = "system", "selected", "")%> >������</option>
			<option value="etc"		<%= Chkiif(lastUserid = "etc", "selected", "")%> >������</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="subcmd" value="">
<input type="hidden" name="chgSellYn" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14">
		�˻���� : <b><%= FormatNumber(oOutmall.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oOutmall.FTotalPage,0) %></b>
	</td>
	<td align="right" valign="top">
		���û�ǰ�� ǰ����
		<input class="button" type="button" id="btnSellYn" value="����" onClick="etcmallSellYnProcess('N', '<%= mallid %>');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td>������</td>
	<td>API�׼�</td>
	<td>�ƿ����ڵ�</td>
	<td>��ǰ�ڵ�</td>
	<td>�ɼ��ڵ�</td>
	<td>�켱����</td>
	<td>��Ͻð�</td>
	<td>ť�����ð�</td>
	<td>API�Ϸ�ð�</td>
	<td>�����Ǹ�</td>
	<td>���м�</td>
	<td>��������</td>
	<td>����ID</td>
	<td width="300">Message</td>
</tr>
<% For i = 0 To oOutmall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oOutmall.FItemList(i).FMidx %>"></td>
	<td><%= oOutmall.FItemlist(i).FMallid %></td>
	<td><%= oOutmall.FItemlist(i).FApiAction %></td>
	<td><%= oOutmall.FItemlist(i).FOutmallGoodno %></td>
	<td><%= oOutmall.FItemlist(i).FItemid %></td>
	<td><%= oOutmall.FItemlist(i).FItemoption %></td>
	<td><%= oOutmall.FItemlist(i).FPriority %></td>
	<td><%= oOutmall.FItemlist(i).FRegdate %></td>
	<td><%= oOutmall.FItemlist(i).FReaddate %></td>
	<td><%= oOutmall.FItemlist(i).FFindate %></td>
	<td>
		<%
			If oOutmall.FItemlist(i).FGSShopSellyn = "Y" Then
				response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			Else
				response.write "<font color='RED'>"&oOutmall.FItemlist(i).FGSShopSellyn&"</font>"
			End If
		%>
	</td>
	<td><%= oOutmall.FItemlist(i).FAccFailCnt %></td>
	<td>
	<%
		Select Case oOutmall.FItemlist(i).FResultCode
			Case "OK"		response.write "<font color='BLUE'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case "ERR"		response.write "<font color='RED'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
			Case Else		response.write "<font color='GRAY'>"&oOutmall.FItemlist(i).FResultCode&"</font>"
		End Select
	%>
	</td>
	<td><%= oOutmall.FItemlist(i).FLastUserid %></td>
	<td width="300"><%= oOutmall.FItemlist(i).FLastErrMsg %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	<% If oOutmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oOutmall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oOutmall.StartScrollPage To oOutmall.FScrollCount + oOutmall.StartScrollPage - 1 %>
		<% If i>oOutmall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oOutmall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<% Set oOutmall = Nothing %>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="400"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
