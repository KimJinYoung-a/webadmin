<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/pop_dealitem_copy.asp
' Description :  �� ��ǰ ���� ����Ʈ
' History : 2023.01.10 ������
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
	'��������
	Dim iCurrpage, iPageSize, iPerCnt, isResearch, sSdate, sEdate, intLoop, stext, dispCate, idx
	Dim oDeal, arrList, iTotCnt, iTotalPage, strTxt, sdiv, datediv, viewdiv, isusing, arrCate, maxDepth

	idx = requestCheckVar(Request("idx"),10)	'���� ������ ��ȣ
	dispCate	= requestCheckVar(Request("disp"),16) 		'���� ī�װ�
	maxDepth = 2
	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����

	isusing 		= requestCheckVar(Request("isusing"),1)
	viewdiv 		= requestCheckVar(Request("viewdiv"),1)
	datediv 		= requestCheckVar(Request("datediv"),1)
	sdiv 		= requestCheckVar(Request("sdiv"),10)
	strTxt 		= requestCheckVar(Request("stext"),32)
	
	isResearch = requestCheckVar(Request("isResearch"),1)
	if isResearch ="" then isResearch ="0"
	'## �˻� #############################
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	'������ ��������
	set oDeal = new ClsDeal
		oDeal.FCPage = iCurrpage		'����������
		oDeal.FPSize = iPageSize		'���������� ���̴� ���ڵ尹��
		oDeal.FSearchDateDiv 	= datediv	'�˻��� ����
		oDeal.FSsDate 	= sSdate	'�˻� ������
		oDeal.FSeDate 	= sEdate	'�˻� ������
		oDeal.FSearchDiv 	= sdiv	'�˻�����
		oDeal.FSeTxt 	= strTxt	'�˻���
		oDeal.FSViewDiv 	= viewdiv	'���� ����
		oDeal.FSIsUsing 	= isusing	'��� ����
		oDeal.FSdispCate 	= dispCate	'����ī�װ� �˻�
 		arrList = oDeal.fnGetCopyItemDealList	'�����͸�� ��������
 		iTotCnt = oDeal.FTotCnt	'��ü ������  ��
 	set oDeal = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	function jsSearch(sType){
		var frm = document.frmEvt
		if (sType == "A"){
				frm.iSD.value = "";
				frm.iED.value = "";
				frm.eventstate.value = "";
				frm.sEtxt.value = "";
				frm.selC.value = "";
		}
		if(frm.sdiv.value=="itemid" && frm.stext.value!=""){
			if(isNaN(frm.stext.value)){
				alert("��ǰ��ȣ �˻��� ���ڸ� �Է����ּ���!");
				return false;
			}
		}

		frm.submit();
	}

	function jsCopyItem(dealcode){
        if(confirm("��ǰ�� ���� �Ͻðڽ��ϱ�?")){
            $.ajax({
                type: "POST",
                url: "dodealitemcopy.asp",
                data: "mode=copy&idx=<%=idx%>&dealcode="+dealcode,
                cache: false,
                async: false,
                success: function(data) {
                    if(data.response=="ok") {
                        alert(data.message);
						opener.jsItemCopyAfter();
						self.close();
                    } else {
                        alert(data.message);
                    }
                }
            });
        }
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEvt" method="get" onSubmit="return jsSearch('E');">
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">�˻� ����</td>
	<td>
		<table>
		<tr>
			<td>
				�˻��� : 
				<select name="sdiv" class="select">
					<option value="itemid"<% If sdiv="itemid" Then Response.write " selected" %>>����ǰ�ڵ�</option>
					<option value="itemname"<% If sdiv="itemname" Then Response.write " selected" %>>��ǰ��</option>
					<option value="register"<% If sdiv="register" Then Response.write " selected" %>>�ۼ���</option>
					<option value="makerid"<% If sdiv="makerid" Then Response.write " selected" %>>�귣����̵�</option>
				</select>
				<input type="text" name="stext" size="50" value="<%=strTxt%>" onkeydown="if(event.keyCode==13) jsSearch('E');">
			</td>
		</tr>
		</table>
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" align="center"><input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch('E');"></td>
</tr>
</form>
</table><br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="13">
			<table width="100%">
			<tr>
				<td>�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>����ǰ�ڵ�</td>
		<td>ī�װ�</td>
		<td>����</td>
		<td>����ǰ��</td>
		<td>��ǰ��</td>
	 </tr>
	 <% If isArray(arrList) Then %>
	 <% For intLoop = 0 To UBound(arrList,2) %>
	 <tr bgcolor="#FFFFFF" onclick="jsCopyItem(<%=arrList(0,intLoop)%>)">
		<td align="center"><%=arrList(1,intLoop)%></td>
		<td align="center">
		<%
			If arrList(12,intLoop) <> "" Then
			arrCate = Split(arrList(12,intLoop),"^^")
			If ubound(arrCate)>0 Then
			Response.write arrCate(0) & " > " & arrCate(1)
			Else
			Response.write arrCate(0)
			End If
			End If
		%>
		</td>
		<td align="center"><% If arrList(2,intLoop)="1" Then %>��õ�<% Else %>�Ⱓ��<% End If %></td>
		<td><%=arrList(5,intLoop)%></td>
		<td align="center">
			<%=arrList(15,intLoop)%>
		</td>
	 </tr>
	 <% Next %>
	 <% Else %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" align="center" height="25">
			��ϵ� ������ �����ϴ�.
		</td>
	 </tr>
	 <% End If %>
	 <tr bgcolor="#FFFFFF">
		<td colspan="13" bgcolor="#FFFFFF" align="center">
			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,"" %>
		</td>
	 </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->