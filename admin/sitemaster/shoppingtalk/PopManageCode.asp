<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : GIFT TALK �ڵ� ����
' Hieditor : ���ر� ����
'			 2022.07.08 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/sitemaster/shoppingtalk/classes/shoppingtalkCls.asp" -->

<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>

<%
	Dim cTalkCode, vDepth, vCode, arrList, intLoop, vKeyword
	Dim iCodeValue, sCodeDesc, iCodeSort, blnUsing, selCodeType
	vKeyword = Request("keyword1")
	vCode = Request("code")
	
	If vKeyword = "" Then
		vDepth = "1"
	Else
		vDepth = "2"
	End If
	
	SET cTalkCode = New CShoppingTalk
	If vCode <> "" Then
		cTalkCode.FRectCode = vCode
		cTalkCode.fnShoppingTalkCodeDetail
		iCodeValue = cTalkCode.FOneItem.FCode
		sCodeDesc = ReplaceBracket(cTalkCode.FOneItem.FCodename)
		iCodeSort = cTalkCode.FOneItem.FSortNo
		blnUsing = cTalkCode.FOneItem.FUseYN
	Else
		iCodeSort = "99"
		blnUsing = "y"
	End If
	
	cTalkCode.FPageSize = 100
	cTalkCode.FCurrpage = 1
	cTalkCode.FRectDepth = vDepth
	cTalkCode.FRectCode = vKeyword
	'cTalkCode.FRectUseYN = "y"
	arrList = cTalkCode.fnShoppingTalkCodeList
%>
<script type='text/javascript'>
<!--
	// �ڵ�Ÿ�� �����̵�
	function jsSetCode(iCodeValue){	
		self.location.href = "PopManageCode.asp?keyword1="+iCodeValue+"";
	}
	
	function jsUpdateCode(keyword,Code){	
		self.location.href = "PopManageCode.asp?keyword1="+keyword+"&Code="+Code+"";
	}
	
	//�ڵ� �˻�
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	//�ڵ� ���
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.NewCode.value) {
			alert("�ڵ尪�� �Է��� �ּ���");
			frm.NewCode.focus();
			return false;
		}
			 
		if(!frm.NewCodename.value) {
			alert("�ڵ���� �Է��� �ּ���");
			frm.NewCodename.focus();
			return false;
		}

		return true;
	}
	
//-->
</script>
<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->	
		<form name="frmReg" method="post" action="procCode.asp" onSubmit="return jsRegCode();">	
		<input type="hidden" name="depth" value="<%=vDepth%>">
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<tr>			
			<td>	+ �ڵ� ��� �� ����</td>
		</tr>	
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">										
				<% If vKeyword <> "" Then %>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�ڵ�Ÿ��</td>
					<td bgcolor="#FFFFFF">
						<select name="NewKeyword1">
						<%= keywordSelectBox (vKeyword, iCodeValue)%>
						</select>
					</td>
				</tr>
				<% End If %>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�ڵ尪</td>
					<td bgcolor="#FFFFFF"><%IF iCodeValue ="" THEN%><input type="text" size="4" maxlength="4" name="NewCode" value="<%=vKeyword%>">
						<%ELSE%><%=iCodeValue%><input type="hidden" size="4" maxlength="4" name="NewCode" value="<%=iCodeValue%>">
						<%END IF%>
						* �⺻�� : <font color=blue><b>�빮�ھ��ĺ�(2�ڸ�)</b></font>
					</td>
				</tr>					
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ڵ��</td>
					<td bgcolor="#FFFFFF"><input type="text" size="15" maxlength="16" name="NewCodename" value="<%= ReplaceBracket(sCodeDesc) %>"> * ' �Ǵ� " �� �ȵ˴ϴ�.</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ڵ� ���ļ���</td>
					<td bgcolor="#FFFFFF"><input type="text" size="4" maxlength="10" name="NewSort" value="<%=iCodeSort%>"> * ���ڰ� �������� ��ܿ� �ֽ��ϴ�.</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">��뿩��</td>
					<td bgcolor="#FFFFFF">
					<input type="radio" value="y" name="NewUseyn" <%IF blnUsing ="y" THEN%>checked<%END IF%>>��� 
					<input type="radio" value="n" name="NewUseyn" <%IF blnUsing ="n" THEN%>checked<%END IF%>>������ </td>
				</tr>
				</table>		
			</td>
		</tr>
		<tr>
			<td align="right"><input type="image" src="/images/icon_save.gif"> 
				<a href="javascript:jsSetCode('','<%=selCodeType%>')"><img src="/images/icon_cancel.gif" border="0"></a></td>
		</tr>	
		<tr>
			<td colspan="2"><hr width="100%"></td>
		</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<form name="frmSearch" method="post" action="PopManageCode.asp">
	<td colspan="2">+ �ڵ� ����Ʈ</td>
</tr>
<tr>
	<td>�ڵ�Ÿ�� : 
		<select name="keyword1" onChange="jsSetCode(this.value);">
		<%= keywordSelectBox (vKeyword, "")%>
		</select>
	</td>
	<td align="right"><a href="javascript:jsSetCode('','<%=selCodeType%>');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">	
		<div id="divList" style="height:410px;overflow-y:scroll;">	
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">				
		<tr bgcolor="#EFEFEF">			
			<td  align="center" width="50">�ڵ尪</td>
			<td  align="center">�ڵ��</td>
			<td  align="center">���ļ���</td>
			<td  align="center">��뿩��</td>
			<td  align="center">ó��</td>
		</tr>
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
		<tr bgcolor="#FFFFFF">			
			<td  align="center"><%=arrList(1,intLoop)%></td>
			<td  align="center"><%=arrList(2,intLoop)%></td>
			<td  align="center"><%=arrList(3,intLoop)%></td>
			<td  align="center"><%=arrList(4,intLoop)%></td>
			<td  align="center">
				<input type="button" value="����" onClick="javascript:jsUpdateCode('<%=vKeyword%>','<%=arrList(1,intLoop)%>');" class="input_b">				
			</td>
		</tr>
			<%Next%>
		<%ELSE%>	
		<tr bgcolor="#FFFFFF">			
			<td colspan="5" align="center">��ϵ� ������ �����ϴ�.</td>
		</tr>	
		<%End if%>		
		</table>
		</div>
	</td>
	</form>
</tr>
</table>
<% SET cTalkCode = Nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->