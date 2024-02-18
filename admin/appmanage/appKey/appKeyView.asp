<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/appKeyCls.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->

<%
	Dim cOneAppKey, idx, vType, vOsType, vAppVersion, vValidationKey, vRegDate, vLastUpDate, vAdminId, vAdminName, vIsUsing
	idx	= getNumeric(requestCheckVar(request("idx"),10))

	if idx<>"" then
		SET cOneAppKey = New CappKey
		cOneAppKey.FRectIdx = idx
		cOneAppKey.GetOneAppKey

		if cOneAppKey.FResultCount>0 then
			vType			= cOneAppKey.FOneKey.Ftype
			vOsType			= cOneAppKey.FOneKey.FosType
			vAppVersion		= cOneAppKey.FOneKey.FappVersion
			vValidationKey	= cOneAppKey.FOneKey.FvalidationKey
			vRegDate		= cOneAppKey.FOneKey.FregDate
			vLastUpDate		= cOneAppKey.FOneKey.FlastUpDate
			vAdminId		= cOneAppKey.FOneKey.FadminId
			vAdminName		= cOneAppKey.FOneKey.FadminName
			vIsUsing		= cOneAppKey.FOneKey.FisUsing
		end if

		SET cOneAppKey = Nothing
	end if
%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function fnChgLinkType(val) {
	switch(val) {
		case "event":
			document.frm1.linkTitle.value = "�̺�Ʈ";
			document.frm1.linkURL.value = "/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�";
			break;
		case "spevt":
			document.frm1.linkTitle.value = "��ȹ��";
			document.frm1.linkURL.value = "/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�";
			break;
		case "prd":
			document.frm1.linkTitle.value = "��ǰ����";
			document.frm1.linkURL.value = "/category/category_itemprd.asp?itemid=��ǰ�ڵ�";
			break;
		default:
			document.frm1.linkTitle.value = "";
			document.frm1.linkURL.value = "";
	}
}

// ����� Ȯ�� �� ó��
function fnSubmit(frm) {
	if(!frm.appversion.value) {
		alert("�۹����� �Է����ּ���.");
		frm.appversion.focus();
		return false;
	}

	if(!frm.validationkey.value) {
		alert("����Ű�� �Է����ּ���.");
		frm.validationkey.focus();		
		return false;
	}

	if(confirm("�Է��Ͻ� �������� ����Ͻðڽ��ϱ�?")){
		frm.submit();
	}

}
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="â�ݱ�" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl" style="padding-top:20px;">
		<div class="contTit bgNone">
			<h2>APP����Ű ���</h2>
		</div>
		<div class="cont">
			<form name="frm1" action="doAppKeyReg.asp" method="post" style="margin:0px;">
			<input type="hidden" name="idx" value="<%=idx%>">
			<input type="hidden" name="mode" value="<%=chkiif(idx="" or isNull(idx),"add","modi")%>">
				<table class="tbType1 writeTb" bgcolor="#FFFFFF">
					<tbody>
						<tr>
							<th>�۱���</th>
							<td height="30" style="padding-left:5px;">
								<select name="type" class="formSlt" >
									<option value="wishapp" <%=chkIIF(vType="wishapp" or vType="","selected","")%>>���þ�</option>
									<option value="hitchhiker" <%=chkIIF(vType="hitchhiker","selected","")%>>��ġ����Ŀ</option>
								</select>
							</td>
						</tr>
						<tr>
							<th>OS����</th>
							<td height="30" style="padding-left:5px;">
								<select name="ostype" class="formSlt" >
									<option value="ios" <%=chkIIF(vOsType="ios" or vOsType="","selected","")%>>iOS</option>
									<option value="android" <%=chkIIF(vOsType="android","selected","")%>>Android</option>
								</select>
							</td>
						</tr>						
						<tr>
							<th>�۹���</th>
							<td height="30" style="padding-left:5px;">
								<input type="text" name="appversion" value="<%=vAppVersion%>" class="formTxt" size="50" maxlength="100" />
							</td>
						</tr>
						<tr>
							<th>����Ű</th>
							<td height="30" style="padding-left:5px;">
								<input type="text" name="validationkey" value="<%=vValidationKey%>" class="formTxt" size="50" maxlength="100" />
							</td>
						</tr>						
						<tr>
							<th>��뿩��</th>
							<td height="30" style="padding-left:5px;">
								<label><input type="radio" name="isUsing" value="Y" class="formCheck" <%=chkIIF(vIsUsing="" or vIsUsing="Y","checked","")%> /> ���</label>
								<label><input type="radio" name="isUsing" value="N" class="formCheck" <%=chkIIF(vIsUsing="N","checked","")%> /> ������</label>
							</td>
						</tr>
					</tboby>
				</table>

				<div class="tPad15 ct">
					<input type="button" value="�� ��" onclick="if(confirm('�۾��� ����ϰ� â�� �ݰڽ��ϱ�?')){self.close();}" class="btn3 btnDkGy" style="margin-right:30px;" />
					<input type="button" value="�� ��" onclick="fnSubmit(this.form);" class="btn3 btnRd" />
				</div>
			</form>
		</div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->