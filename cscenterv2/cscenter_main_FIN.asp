<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.10.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim sqlStr

'// ============================================================================
' ���̵� Ȯ�λ���
Dim csMainUserID, csMainUserName
csMainUserID	= req("csMainUserID", session("ssBctId") )
csMainUserName	= session("ssBctCname")


'// ============================================================================
'��ó���޸�
dim CSMemoNotFinishFIN

sqlStr = " [db_academy].[dbo].[usp_ACA_GetMiFinishMemo] '" + CStr(csMainUserID) + "' "

rsACADEMYget.CursorLocation = adUseClient
rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
if  not rsACADEMYget.EOF  then
    CSMemoNotFinishFIN = rsACADEMYget("CSMemoNotFinish")
end if
rsACADEMYget.close

%>
<style>
//
</style>
<script language="JavaScript" src="/cscenterv2/js/cscenter.js?v=1.1"></script>
<script language="javascript">
function jsReload() {
	var filename = window.location.href.split("/").pop();
	if (filename.split("?").length > 1) {
		filename = filename.split("?")[0];
	}

	location.href = filename + "?menupos=" + document.frm.menupos.value + "&csMainUserID=" +  document.getElementById('csMainUserID').value;
}

// orderserial, userid, finishyn, writeUser
function cscenter_memo_list_FIN(args) {
	var params = object2queryparams(args);
	var popwin = window.open("/cscenterv2/history/history_memo_list.asp" + params,"cscenter_memo_list_FIN","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="post" action="cscenter_main_process.asp">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="mode" value="">
		<input type="hidden" name="csTime" value="">
	</form>
	<tr>
		<!-- ���ʸ޴� ���� -->
		<td width="33%" valign="top">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr valign="top">
					<td>
        				<!-- �ֹ������˻� -->
						<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        					<tr bgcolor="<%= adminColor("menubar") %>">
        						<td>
        							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
										<tr height="25">
            								<td>
            			    					<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����� �˻�</b>
            								</td>
            								<td align="right">
            									<a href="#" onclick="popCallRingFingers({sitename:'academy'}); return false;"> ���� ��û�������� <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
												<a href="#" onclick="popCallRingFingers({sitename:'diyitem'}); return false;"> DIY �ֹ��������� <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            								</td>
            							</tr>
            						</table>
            					</td>
            				</tr>
            			</table>
        				<!-- �ֹ������˻� -->
        			</td>
				</tr>
				<tr valign="top">
					<td height="10"></td>
				</tr>
			</table>
		</td>
		<!-- ���ʸ޴� �� -->
		<td width="10"></td>
		<!-- ����޴� ���� -->
		<td width="33%" valign="top">

		</td>
		<!-- ����޴� �� -->
		<td width="10"></td>
		<!-- �����ʸ޴� ���� -->
		<td valign="top">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr valign="top">
					<td>
						<!-- ���ΰ�ħ ���� -->
						<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        					<tr bgcolor="<%= adminColor("tabletop") %>">
        						<td>
        							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
										<tr height="25">
                        					<td>
            			    					<img src="/images/icon_star.gif" align="absbottom">
												<b>ID : </b>
												<input type="text" class="text" id="csMainUserID" value="<%=csMainUserID%>" size="10">
												<input type="button" class="button" value="�˻�" onclick="jsReload();">
												<!-- �ʱ�α��ν� �α��� ���̵�� ���� / �ٸ����̵�ε� �˻������ϵ��� -->
            								</td>
            								<td align="right">
            			    					<a href="#" onclick="document.location.reload(); return false;">
        											���ΰ�ħ
        											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        										</a>
            								</td>
            							</tr>
            						</table>
            					</td>
            				</tr>
            			</table>
            			<!-- ���ΰ�ħ �� -->
					</td>
				</tr>
				<tr valign="top">
					<td height="10"></td>
				</tr>
				<tr valign="top">
					<td>
						<!-- ���̵� Ȯ�λ��� -->
						<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        					<tr bgcolor="<%= adminColor("tabletop") %>">
        						<td>
        							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            							<tr height="25">
            								<td style="border-bottom:1px solid #BABABA">
            									<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���̵� Ȯ�λ���</b>
            								</td>
            								<td align="right" style="border-bottom:1px solid #BABABA">
            									&nbsp;
            								</td>
            							</tr>
										<tr height="25">
            								<td>�ΰŽ� ��ó���޸�</td>
            								<td align="right">
            									<b><%= CSMemoNotFinishFIN %></b> ��
        				    					<a href="#" onclick="cscenter_memo_list_FIN({writeUser:'<%=csMainUserID%>', finishyn:'N'}); return false;">
                    								<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    							</a>
            								</td>
            							</tr>
            						</table>
            					</td>
            				</tr>
            			</table>
        				<!--  ���̵� Ȯ�λ��� ��-->
					</td>
				</tr>
				<tr valign="top">
					<td height="10"></td>
				</tr>
				<tr valign="top">
					<td>
            			<!-- SMS MAIL -->
						<table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        					<tr bgcolor="<%= adminColor("tabletop") %>">
        						<td>
        							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
										<tr height="25">
            								<td>
            			    					<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
													<tr height="20">
		            									<td align="center">
		            			    						<a href="" onclick="PopCSSMSSendNew({}); return false;">SMS�߼�</a>
		            									</td>
		            								</tr>
		            							</table>
            								</td>
            								<td width="5"></td>
            								<td>
            			    					<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
													<tr height="20">
		            									<td align="center">
															<!--
		            			    						<a href="javascript:PopCSMailSend({});">���Ϲ߼�</a>
															-->
															<a href="javascript:PopCSMailSend('','');">���Ϲ߼�</a>
		            									</td>
		            								</tr>
		            							</table>
            								</td>
            							</tr>
            						</table>
            					</td>
            				</tr>
            			</table>
        				<!-- SMS MAIL -->
					</td>
				</tr>
			</table>
		</td>
		<!-- �����ʸ޴� �� -->
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
