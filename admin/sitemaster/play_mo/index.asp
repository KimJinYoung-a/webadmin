<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �÷��̸����
' Hieditor : ����ȭ ����
'			 2022.07.07 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
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
<link rel="stylesheet" type="text/css" href="/css/adminPartnerDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script type='text/javascript'>

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
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->
<%
	Dim i, cPlay, vPage, vIsUsing, vType, vTitle, vPartMDID, vPartWDID, vPartPBID, vState
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vIsUsing = requestCheckVar(request("isusing"),1)
	vType = requestCheckVar(request("playtype"),2)
	vTitle = requestCheckVar(request("title"),200)
	vState = requestCheckVar(request("state"),2)
	vPartMDID = NullFillWith(requestCheckVar(request("partmdid"),50),"")
	vPartWDID = NullFillWith(requestCheckVar(request("partwdid"),50),"")
	vPartPBID = NullFillWith(requestCheckVar(request("partpbid"),50),"")
	
	SET cPlay = New CPlayMoContents
	cPlay.FCurrPage = vPage
	cPlay.FPageSize = 20
	cPlay.FRectIsusing = vIsUsing
	cPlay.FRectType = vType
	cPlay.FRectTitle = vTitle
	cPlay.FRectState = vState
	cPlay.FRectMDID = vPartMDID
	cPlay.FRectWDID = vPartWDID
	cPlay.FRectPBID = vPartPBID
	cPlay.fnPlayMoList
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goNewReg(idx){
	var winPlay;
	winPlay = window.open('write.asp?idx='+idx,'winPlay','width=1400, height=800, scrollbars=yes');
	winPlay.focus();
}
function goPlayType(idx){
	var winPlayType;
	winPlayType = window.open('pop_type.asp','winPlayType','width=410, height=570');
	winPlayType.focus();
}
function goStyleCode(idx){
	var winStyleCode;
	winStyleCode = window.open('pop_style.asp','winStyleCode','width=410, height=570');
	winStyleCode.focus();
}
function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}
//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
function jsImgView(sImgUrl){
 var wImgView;
 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
 wImgView.focus();
}
function jsTagview(idx,type){	
	var poptagm;
	poptagm = window.open('pop_tagReg.asp?idx='+idx+'&playcate='+type+'','poptagm','width=500,height=400,scrollbars=yes,resizable=yes');
	poptagm.focus();
}
function jsItem(idx,type){
	var popPItem;
	popPItem = window.open('item.asp?playidx='+idx+'&playcate='+type+'','popPItem','width=1200,height=1000,scrollbars=yes,resizable=yes');
	popPItem.focus();
}
</script>

<div class="contSectFix scrl">
	<!-- ��� �˻��� ���� -->
	<form name="frm1" method="get" action="" style="margin:0px;">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="term1">��� ���� :</label>
					<select name="isusing" class="formSlt">
						<option value=""> - ���� - </option>
						<option value="Y" <%=CHKIIF(vIsUsing="Y","selected","")%>>Y(�����)</option>
						<option value="N" <%=CHKIIF(vIsUsing="N","selected","")%>>N(������)</option>
					</select>
				</li>
				<li>
					<label class="formTit" for="term1">�� �� :</label>
					<select name="state" class="formSlt">
						<%=fnStateSelectBox("select",vState)%>
					</select>
				</li>
				<li>
					<label class="formTit" for="term1">�� �� :</label>
					<select name="playtype" class="formSlt">
						<%=fnTypeSelectBox("select",vType,"Y")%>
					</select>
				</li>
				<li>
					<label class="formTit" for="term1">�� �� :</label>
					<input type="text" class="formTxt" name="title" value="<%=vTitle%>" style="width:200px" />
				</li>
				<li>
					<label class="formTit" for="term1">����� :</label>
					<% sbGetpartid "partmdid",vPartMDID,"","11,14,21,22,23" %>
				</li>
				<li>
					<label class="formTit" for="term1">WD :</label>
					<% sbGetpartid "partwdid",vPartWDID,"","12" %>
				</li>
				<li>
					<label class="formTit" for="term1">�ۺ��� :</label>
					<select name="partpbid">
						<option value="">����</option>
						<option value="happyngirl" <%=CHKIIF(vPartPBID="happyngirl","selected","")%>>�ּ���</option>
						<option value="kyungae13" <%=CHKIIF(vPartPBID="kyungae13","selected","")%>>�����</option>
						<option value="jinyeonmi" <%=CHKIIF(vPartPBID="jinyeonmi","selected","")%>>������</option>
					</select>
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="�˻�" />
	</div>
	</form>
	
	<div class="pad20">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* �� <%=cPlay.FTotalCount%> �� / idx, ������ ���� Sorting �Ǿ��ֽ��ϴ�.</p>
			</div>
			<div class="ftRt">
				<p class="btn2 cBk1 ftLt"><a href="javascript:goStyleCode('');"><span class="eIcon"><em class="fIcon">��Ÿ���ڵ����</em></span></a></p>
				<p class="ftLt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
				<p class="btn2 cBk1 ftLt"><a href="javascript:goPlayType('');"><span class="eIcon"><em class="fIcon">�з�����</em></span></a></p>
				<p class="ftLt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
				<p class="btn2 cBk1 ftLt"><a href="javascript:goNewReg('');"><span class="eIcon"><em class="fIcon">�űԵ��</em></span></a></p>
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>idx</div></th>
					<th><div>No | �ؽ�Ʈ</div></th>
					<th><div>�з�</div></th>
					<th><div>�̹���</div></th>
					<th><div>����</div></th>
					<th><div>��뿩��</div></th>
					<th><div>�켱������ȣ</div></th>
					<th><div>��ȸ��(�α��)</div></th>
					<th><div>������</div></th>
					<th><div>�۾�����</div></th>
					<th><div>�����</div></th>
					<th><div>�۾���</div></th>
					<th><div>���</div></th>
				</tr>
				</thead>
				<tbody>
				<%
					If cPlay.FResultCount > 0 Then
						For i=0 To cPlay.FResultCount-1
				%>
						<tr>
							<td><%=cPlay.FItemList(i).Fidx%></td>
							<td><%=cPlay.FItemList(i).Fviewno%> | <%=cPlay.FItemList(i).Fviewnotxt%></td>
							<td><%=cPlay.FItemList(i).Ftypename%></td>
							<td><img src="<%=cPlay.FItemList(i).Flistimg%>" height="100" style="cursor:pointer;" onclick="jsImgView('<%=cPlay.FItemList(i).Flistimg%>');"></td>
							<td>
								<%= ReplaceBracket(cPlay.FItemList(i).Ftitle) %>
								<br /><br /><%=fnStateSelectBox("one",cPlay.FItemList(i).Fstate)%>
							</td>
							<td><%=cPlay.FItemList(i).Fisusing%></td>
							<td><%=cPlay.FItemList(i).Fsortno%></td>
							<td><%=cPlay.FItemList(i).Ffavcnt%></td>
							<td><%=cPlay.FItemList(i).Fstartdate%></td>
							<td>
								����� : <%=cPlay.FItemList(i).Fregdate%><br /><br />
								���� : <%=cPlay.FItemList(i).Flastadminid%><br /><%=cPlay.FItemList(i).Flastupdate%></td>
							<td><%=cPlay.FItemList(i).FpartMDname%></td>
							<td>
								WD:<%=cPlay.FItemList(i).FpartWDname%><br />
								PB:<%=cPlay.FItemList(i).FpartPBname%>
							</td>
							<td>
								<input type="button" onClick="goNewReg('<%=cPlay.FItemList(i).Fidx%>');" value="�� ��"><br /><br />
								<input type="button" onClick="jsTagview('<%=cPlay.FItemList(i).Fidx%>','<%=cPlay.FItemList(i).Ftype%>');" value="�� ��"><br /><br />
								<input type="button" onClick="jsItem('<%=cPlay.FItemList(i).Fidx%>','<%=cPlay.FItemList(i).Ftype%>');" value="�� ǰ">
							</td>
						</tr>
				<%
						Next
					End If
				%>
				</tbody>
			</table>
			<div class="ct tPad20 cBk1">
				<% if cPlay.HasPreScroll then %>
				<a href="javascript:searchFrm('<%= cPlay.StartScrollPage-1 %>')">[pre]</a>
				<% else %>
	    			[pre]
	    		<% end if %>
	    		
	    		<% for i=0 + cPlay.StartScrollPage to cPlay.FScrollCount + cPlay.StartScrollPage - 1 %>
	    			<% if i>cPlay.FTotalpage then Exit for %>
	    			<% if CStr(vPage)=CStr(i) then %>
	    			<span class="cRd1">[<%= i %>]</span>
	    			<% else %>
	    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
	    			<% end if %>
	    		<% next %>
				
				<% if cPlay.HasNextScroll then %>
	    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
	    		<% else %>
	    			[next]
	    		<% end if %>
			</div>
		</div>
	</div>
</div>

<% SET cPlay = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->