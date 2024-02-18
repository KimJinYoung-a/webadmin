<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : �ý��۰��� > VPN������Ȳ
' History : ������ ����
'			2017.05.19 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = requestCheckvar(request("menupos"),10)
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// ���ã��
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;

	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
<% if session("sslgnMethod")<>"S" then %>
<!-- USBŰ ó�� ���� (2008.06.23;������) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USBŰ ó�� �� -->
<% end if %>
</head>
<body <% if session("sslgnMethod")<>"S" then %>onload="checkUSBKey()"<% end if %>>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cooperate/vpnconnectCls.asp" -->
<%
	Dim i, cVPN, vPage, sTime1, sTime2, eTime1, eTime2, vIsSign, vUserID
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	sTime1 = requestCheckVar(request("sTime1"),10)
	sTime2 = requestCheckVar(request("sTime2"),10)
	eTime1 = requestCheckVar(request("eTime1"),10)
	eTime2 = requestCheckVar(request("eTime2"),10)
	vIsSign = requestCheckVar(request("issign"),1)
	vUserID = requestCheckVar(request("userid"),50)
	
	SET cVPN = New Cvpnconnect_list
	cVPN.FCurrPage = vPage
	cVPN.FPageSize = 25
	cVPN.FRectSTime1 = sTime1
	cVPN.FRectSTime2 = sTime2
	cVPN.FRectETime1 = eTime1
	cVPN.FRectETime2 = eTime2
	cVPN.FRectIsSign = vIsSign
	cVPN.FRectUserID = vUserID
	cVPN.sbVPNLogList
%>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goInsertLog(idx){
	var winLog;
	winLog = window.open('write.asp?idx='+idx,'winLog','width=1400, height=600');
	winLog.focus();
}
function goInsertWhycon(idx){
	var winWhy;
	winWhy = window.open('whycon.asp?idx='+idx,'winWhy','width=1400, height=600');
	winWhy.focus();
}
function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}
function jsSign(idx){
    
	document.signfrm.idx.value = idx;
	document.signfrm.gubun.value = 'sign';
	document.signfrm.submit();
}
function jsonedel(idx){
	<% If not(session("ssBctId") = "tozzinet" or session("ssBctId") = "kei0329" or session("ssBctId") = "coolhas") Then %>
		alert('[�����ڱ���]���� �Ұ�');
		return;
	<% end if %>

	document.signfrm.idx.value = idx;
	document.signfrm.gubun.value = 'onedel';
	document.signfrm.submit();
}
function jsShowDiv(idx){
	$("#div"+idx+"").show();
}
function jsHideDiv(idx){
	$("#div"+idx+"").hide();
}
</script>

<div class="contSectFix scrl">
	<!-- ��� �˻��� ���� -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="term1">���ӽð� :</label>
					<input type="text" class="formTxt" id="sTime1" name="sTime1" value="<%=sTime1%>" style="width:100px" placeholder="������" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="sTime1_trigger" alt="�޷����� �˻�" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "sTime1", trigger    : "sTime1_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					~
					<input type="text" class="formTxt" id="sTime2" name="sTime2" value="<%=sTime2%>" style="width:100px" placeholder="������" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="sTime2_trigger" alt="�޷����� �˻�" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "sTime2", trigger    : "sTime2_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
				<!--<li>
					<label class="formTit" for="term1">����ð�(���ӽð�����) :</label>
					<input type="text" class="formTxt" id="eTime1" name="eTime1" value="<%=eTime1%>" style="width:100px" placeholder="������" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="eTime1_trigger" alt="�޷����� �˻�" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "eTime1", trigger    : "eTime1_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					~
					<input type="text" class="formTxt" id="eTime2" name="eTime2" value="<%=eTime2%>" style="width:100px" placeholder="������" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="eTime2_trigger" alt="�޷����� �˻�" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "eTime2", trigger    : "eTime2_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>-->
				<li>
					<label class="formTit" for="term1">���� :</label>
					<% drawvpnlistSelectBox "userid", vUserID, "" %>
					<!--<select class="formSlt" id="userid" name="userid" title="���� ����">
						<option value="" <%=CHKIIF(vUserID="","selected","")%>>��ü</option>
						<option value="VPN_kei0329" <%=CHKIIF(vUserID="VPN_kei0329","selected","")%>>VPN_kei0329</option>
						<option value="VPN_kjy8517" <%=CHKIIF(vUserID="VPN_kjy8517","selected","")%>>VPN_kjy8517</option>
						<option value="VPN_kobula" <%=CHKIIF(vUserID="VPN_kobula","selected","")%>>VPN_kobula</option>
						<option value="VPN_thensi" <%=CHKIIF(vUserID="VPN_thensi","selected","")%>>VPN_thensi</option>
						<option value="VPN_tozzinet" <%=CHKIIF(vUserID="VPN_tozzinet","selected","")%>>VPN_tozzinet</option>
						<option value="vpn_tkwon" <%=CHKIIF(vUserID="vpn_tkwon","selected","")%>>vpn_tkwon</option>
						<option value="vpn_corpse2" <%=CHKIIF(vUserID="vpn_corpse2","selected","")%>>vpn_corpse2</option>
					</select>-->
				</li>
				<% If session("ssBctId") = "coolhas" Then %>
				<li>
					<label class="formTit" for="term1">���ο��� :</label>
					<select name="issign" class="select" onChange="frm1.submit();">
						<option value="" <%=CHKIIF(vIsSign="","selected","")%>>-��ü-</option>
						<option value="x" <%=CHKIIF(vIsSign="x","selected","")%>>���</option>
						<option value="o" <%=CHKIIF(vIsSign="o","selected","")%>>�Ϸ�</option>
					</select>
				</li>
				<% End If %>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="�˻�" />
	</div>
	</form>
	
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* �� <%=cVPN.FTotalCount%> �� / ���ӽð����� Sorting �Ǿ��ֽ��ϴ�.</p>
			</div>
			<% If C_ADMIN_AUTH or C_OP Then %>
			<div class="ftRt">
				<p class="btn2 cBk1 ftLt"><a href="javascript:goInsertLog('');"><span class="eIcon"><em class="fIcon">�α��Է�</em></span></a></p>
			</div>
			<% End If %>
		</div>
		
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<!--<th><div>idx</div></th>//-->
					<th><div>���ӽð�</div></th>
					<!--<th><div>����ð�</div></th>-->
					<!--<th><div>���ӽð�</div></th>-->
					<th><div>����</div></th>
					<th><div>����IP</div></th>
					<!--<th><div>�Ҵ�IP</div></th>-->
					<!--<th><div>�α��λ���</div></th>-->
					<!--<th><div>���ӻ���</div></th>-->
					<th><div>���ӻ���</div></th>
					<th><div>����</div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cVPN.FResultCount > 0 Then
						For i=0 To cVPN.FResultCount-1
				%>
						<tr>
							<!--<td><%=cVPN.FItemList(i).Fidx%></td>//-->
							<td><%=cVPN.FItemList(i).Fstime%></td>
							<!--<td><%=cVPN.FItemList(i).Fetime%></td>-->
							<!--<td><%=cVPN.FItemList(i).Fequip%></td>-->
							<td><%=cVPN.FItemList(i).Fusername%>(<%=cVPN.FItemList(i).Fuserid%>)</td>
							<td><%=cVPN.FItemList(i).Frealip%></td>
							<!--<td><%'=cVPN.FItemList(i).Fassignip%></td>-->
							<!--<td><%'=cVPN.FItemList(i).Floginstate%></td>-->
							<!--<td><%'=cVPN.FItemList(i).Fconstate%></td>-->
							<td>
								<%
									If isNull(cVPN.FItemList(i).Fwhycon) OR cVPN.FItemList(i).Fwhycon = "" Then 
										If ucase(session("ssBctId")) = Replace(ucase(cVPN.FItemList(i).Fuserid),"VPN_","")Then '2017.02.21 ������ ���� ucase ��ġ ����
											Response.Write "<input type='button' value='�����Է�' onClick='goInsertWhycon(" & cVPN.FItemList(i).Fidx & ");'>"
										Else
											if (session("ssBctId") = "coolhas" and ucase(cVPN.FItemList(i).Fuserid) = ucase("VPN_eastone")) or (session("ssBctId") = "thensi7" and ucase(cVPN.FItemList(i).Fuserid) = ucase("VPN_thensi")) or (session("ssBctId") = "tozzinet" and ucase(cVPN.FItemList(i).Fuserid) = ucase("VPN_tozzinet")) or (session("ssBctId") = "kei0329" and ucase(cVPN.FItemList(i).Fuserid) = ucase("VPN_kei0329")) then
												Response.Write "<input type='button' value='�����Է�' onClick='goInsertWhycon(" & cVPN.FItemList(i).Fidx & ");'>"
											Else
												Response.Write "���� ���Է�"
											End If
										End If
									Else
										Response.Write "<span onClick='jsShowDiv("&cVPN.FItemList(i).Fidx&");' style='cursor:pointer;'>[��������]</span>"
										If session("ssBctId") = Replace(cVPN.FItemList(i).Fuserid,"VPN_","") Then
											Response.Write "[<a href='javascript:goInsertWhycon(" & cVPN.FItemList(i).Fidx & ");'><strong>����</strong></a>]"
										Else
											If (session("ssBctId") = "coolhas" AND cVPN.FItemList(i).Fuserid = "VPN_eastone") Then
												Response.Write "[<a href='javascript:goInsertWhycon(" & cVPN.FItemList(i).Fidx & ");'><strong>����</strong></a>]"
											End If
										End If
										Response.Write "<div id='div"&cVPN.FItemList(i).Fidx&"' onMouseOut='jsHideDiv("&cVPN.FItemList(i).Fidx&");' style='display:none; position:absolute; border:solid 1px #000000; width:200px; padding:3px; background-color:#ffffff;'>"
										Response.Write "�ۼ��� : " & cVPN.FItemList(i).Fwhyuserid & "<br />"
										Response.Write "�ۼ��� : " & cVPN.FItemList(i).Fwhyregdate & "<br />"
										Response.Write Replace(cVPN.FItemList(i).Fwhycon,vbCrLf,"<br />")
										Response.Write "</div>"
									End If
								%>
							</td>
							<td>
								<% If cVPN.FItemList(i).Fsign = "" Then %>
									<% If session("ssBctId") = "coolhas" Then %>
									    <% if NOT (isNull(cVPN.FItemList(i).Fwhycon) OR cVPN.FItemList(i).Fwhycon = "") then %>
										<input type="button" value="����ó��" onClick="jsSign('<%=cVPN.FItemList(i).Fidx%>');">
									    <% end if %>
									<% Else %>
										����ó�����
									<% End If %>
								<% Else %>
									ó����<br /><%=cVPN.FItemList(i).Fsigndate%>
								<% End If %>
								<% if (FALSE) then %>
								<input type="button" value="����" onClick="jsonedel('<%=cVPN.FItemList(i).Fidx%>');" class="button">
							    <% end if %>
							</td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<br />
			<div class="ct tPad20 cBk1">
				<% if cVPN.HasPreScroll then %>
				<a href="javascript:searchFrm('<%= cVPN.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cVPN.StartScrollPage to cVPN.FScrollCount + cVPN.StartScrollPage - 1 %>
	    			<% if i>cVPN.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cVPN.HasNextScroll then %>
	    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>

<% SET cVPN = Nothing %>

<form name="signfrm" action="proc.asp" method="post" target="procfrm" style="margin:0px;">
<input type="hidden" name="gubun" value="">
<input type="hidden" name="idx" value="">
</form>
<iframe name="procfrm" id="procfrm" width="0" height="0" frameborder="0"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->