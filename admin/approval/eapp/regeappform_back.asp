<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ڰ��� �� ����
' History : 2011.03.14 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<!-- #include virtual="/lib/classes/approval/araplinkedmsCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
 Dim iRectMenu
 iRectMenu = "M000"

Dim clsedms, clsALE, arrList, intLoop
Dim icateidx1, icateidx2,sEdmsName
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim iFormType
Dim  sARAPNM,sedmsNM

 	iFormType =  requestCheckvar(Request("rdoT"),1)
 	IF iFormType ="" THEN iFormType = 2

	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

%>
<html>
<head>
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!--
// ������ �̵�
	function jsGoPage(pg)
	{
		document.frm.iCP.value=pg;
		document.frm.submit();
	}

//ī�װ� ����
$(document).ready(function(){
	$("#selC1").change(function(){
		var iValue = $("#selC1").val();
		var url="/admin/approval/edms/ajaxCategory.asp";
		var params = "sMode=CL&ipcidx="+iValue ;

		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){
		 		$("#divCL").html(args);
		 	},
		 	error:function(e){
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 });
	});
});


//���ڰ��� �����ù�� ����
function jsChangeForm(FormType){
	 document.frm.submit();
}

 //���� ����
 function jsSetDoc(edmsidx,isPay){
    isPay="False"; //������ �׳� ���� ǰ�� �� ������û�� �ۼ��� �����׸��� ���� ���� 2013/10/21

 	if(isPay =="False"){	//������û�� �������� ������ �ٷ� ���ڰ��������� �̵�
 		location.href= "regeapp.asp?ieidx="+edmsidx;
 	}else{ //������û�� �����Ǵ� ��� �������� ���� �� �̵�
 	 	var winApp = window.open("/admin/approval/arap_edms/popGetArapEdms.asp?ieidx="+edmsidx,"popApp","width=600, height=600, resizable=yes");
 		winApp.focus();
 	}
 }

 //������ �Ϸ�-> ���������� �̵�
 function jsSelectEApp(iaidx,ieidx){
	location.href= "regeapp.asp?iAidx="+iaidx+"&ieidx="+ieidx;
 }

 //�˻�
 function jsSearch(){
 	document.frm.iC2.value = $("#selC2").val();  //�˻��� ajax ������ �� �Ѱ��ֱ�
 	document.frm.submit();
 }
//-->
</script>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0">
<tr>
	<td valign="top">
		<table width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				<table width="100%"  cellpadding="0" cellspacing="0" class="a" border="0">
				<form name="frm" method="get" action="regeappform.asp">
				<input type="hidden" name="iAidx" value="">
				<input type="hidden" name="ieidx" value="">
				<input type="hidden" name="iC2" value="">
				<input type="hidden" name="iCP" value="1">
				<tr>
					<td>
						<table width="100%"  cellpadding="5"	 cellspacing="1" class="a" border="0" bgcolor="<%= adminColor("tablebg") %>">
						<tr>
							<td width="100"  bgcolor="#DDDDFF" align="center">���ڰ��� ������</td>
							<td  bgcolor="#FFFFFF" >
							   <input type="radio" name="rdoT" value="1" <%IF iFormType = "1" THEN%>checked<%END IF%> onClick="jsChangeForm(1);">��ü����
							   <input type="radio" name="rdoT" value="2" <%IF iFormType = "2" THEN%>checked<%END IF%> onClick="jsChangeForm(2);">�������ù���
							   <input type="radio" name="rdoT" value="3" <%IF iFormType = "3" THEN%>checked<%END IF%> onClick="jsChangeForm(2);">�ֱٻ�빮��
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td style="padding-top:10px;">
					<!-- #include virtual="/admin/approval/eapp/include_regeappform_list.asp" -->
					</td>
				</tr>
				</form>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->