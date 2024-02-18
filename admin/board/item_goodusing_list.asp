<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : ��ǰ�ı� ����
' History	:  ������ ����
'              2021.11.29 �ѿ�� ����(���� �ٿ�ε� �߰�)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/board/lib/classes/itemGoodUsingCls.asp" -->
<%
Dim page, SearchKey1, SearchKey2, selStatus, lp, lp2, sDt, eDt, dispcate, selPoint, chkTerm, srtMethod,blnPhotomode, chkFirst, searchKeyword
Dim strDel, makerid, orderserial
	page 			= requestCheckvar(Request("page"),10)
	SearchKey1 = requestCheckvar(Request("SearchKey1"),32)
	SearchKey2 = requestCheckvar(Request("SearchKey2"),10)
	selStatus = requestCheckvar(Request("selStatus"),1)
	chkTerm 	= requestCheckvar(Request("chkTerm"),10)
	chkFirst 	= requestCheckvar(Request("chkFirst"),2)
	srtMethod = requestCheckvar(Request("srtMethod"),10)
	sDt = requestCheckvar(Request("sDt"),10)
	eDt = requestCheckvar(Request("eDt"),10)
	dispcate = requestCheckvar(Request("disp"),18)
	selPoint = requestCheckvar(Request("selPoint"),10)
	blnPhotomode = requestCheckvar(Request("photomode"),5)
	makerid     = requestCheckvar(request("makerid"),32)
	orderserial = requestCheckvar(request("orderserial"),12)
	searchKeyword = requestCheckvar(request("keyword"),30)

'�⺻�� ����
if page="" then page=1
if selStatus="" then selStatus="Y"
if srtMethod="" then srtMethod="idxDcd"
if sDt="" and chkTerm="" then sDt = date()
if eDt="" and chkTerm="" then eDt = date()

'// ��ǰ �ı� ���
dim oGoodUsing
Set oGoodUsing = new CGoodUsing
	oGoodUsing.FPagesize = 15
	oGoodUsing.FCurrPage = page
	oGoodUsing.FRectSearchKey1 = SearchKey1
	oGoodUsing.FRectSearchKey2 = SearchKey2
	oGoodUsing.FRectselStatus = selStatus
	oGoodUsing.FRectStartDt = sDt
	oGoodUsing.FRectEndDt = eDt
	oGoodUsing.FRectDispcate = dispcate
	oGoodUsing.FRectPoint = selPoint
	oGoodUsing.FRectPhotoMode = blnPhotomode
	oGoodUsing.FRectSort = srtMethod
	oGoodUsing.FRectMakerid = makerid
	oGoodUsing.FRectOrderserial = orderserial
	oGoodUsing.FRectFirst = chkFirst
	oGoodUsing.FRectKeyword = searchKeyword
	oGoodUsing.GetGoodUsingList
%>
<style type="text/css">
.itemBlock {white-space: nowrap;}
</style>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		document.frm.page.value=pg;
		document.frm.action="";
		document.frm.submit();
	}

	// ���� ���� ����
	function chgStatus(v) {
		document.frm.selStatus.value=v;
		document.frm.action="";
		document.frm.submit();
	}

	// ��ǰ�� �˾�
	function viewItemInfo(iid) {
		var PpUp = window.open("<%=wwwurl%>/common/PopZoomItem.asp?itemid="+ iid +"&pop=pop","itemInfo","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,width=720,height=444");
		PpUp.focus();
	}

	// ���Ĺ�� ����
	function ChangeSort(smtd) {
		document.frm.srtMethod.value=smtd;
		document.frm.action="";
		document.frm.submit();
	}

	// ��ü ����,���
	function chgSel_on_off() {
		var frm = document.frm_list;
		if (frm.lineSel.length) {
			for(var i=0;i<frm.lineSel.length;i++) {
				frm.lineSel[i].checked=frm.tt_sel.checked;
			}
		} else {
			frm.lineSel.checked=frm.tt_sel.checked;
		}
	}

	// ��ü�Ⱓ ����
	function swChkTerm(ckt)	{
		if(ckt.checked) {
			frm.sDt.disabled=true;
			frm.eDt.disabled=true;
		} else {
			frm.sDt.disabled=false;
			frm.eDt.disabled=false;
		}
	}

	// ���õ� �׸� ����/����
	function doSubmit(md) {
		var i, chk=0, strMd;
		var frm = document.frm_list;
		if (md=='restore')
			strMd = "����";
		else
			strMd = "����";

		if (frm.lineSel.length) {
			for(i=0;i<frm.lineSel.length;i++) {
				if(frm.lineSel[i].checked) {
					chk++;
				}
			}
		} else {
				if(frm.lineSel.checked) {
					chk++;
				}
		}

		if(chk==0) {
			alert(strMd + "�� ��ǰ�ı⸦ ��� �Ѱ��̻� �������ֽʽÿ�.");
			return;
		} else {
			if(confirm("�����Ͻ� " + chk + "����  �׸��� ��� " + strMd + "�Ͻðڽ��ϱ�?")) {
				frm.mode.value=md;
				frm.action="doItemGoodUsing.asp";
				frm.submit();
			} else {
				return;
			}
		}
	}
	//�̹��� ����
	function showimage(img){
		var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
	}

// ī�װ� ����� ���
function changecontent(){
}

// ������ ��
function fnPointView(frm,sw) {
	if(sw=="on") {
		frm.children(0).style.display="";
	} else {
		frm.children(0).style.display="none";
	}
}

function itemgoodusingExcelDown(){
	var vText = document.frm_list.selDCnt.options[document.frm_list.selDCnt.selectedIndex].text;
	alert('��'+vText+'���� �ٿ�ε� �մϴ�.\n��� ��ٷ� �ֽø� �ٿ�ε尡 �˴ϴ�.');
	document.frm.target = "exceldown";
	document.frm.action = "/admin/board/item_GoodUsing_list_excel.asp"
	document.frm.page.value = document.frm_list.selDCnt.value;
    document.frm.submit();
	document.frm.target = "";
	document.frm.action = ""
}

//-->
</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="srtMethod" value="<%=srtMethod%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left" style="line-height:24px;">
		<span class="itemBlock">ī�װ� <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></span>
		<br />
		<span class="itemBlock">���̵� <input type="text" name="SearchKey1" size="12" value="<%=SearchKey1%>" class="text" /></span>
		<span class="itemBlock">/ ��ǰ��ȣ <input type="text" name="SearchKey2" size="12" value="<%=SearchKey2%>" class="text" /></span>
		<span class="itemBlock">/ �ֹ���ȣ <input type="text" name="orderserial" size="12" value="<%=orderserial%>" class="text" /></span>
		<span class="itemBlock">/ �귣��ID <%	drawSelectBoxDesignerWithName "makerid", makerid %></span>
		<span class="itemBlock">
			/ ���� <select name="selStatus" onchange="chgStatus(this.value)" class="select">
			<option value="A">��ü</option>
			<option value="N">����</option>
			<option value="Y">�Ϲ�</option>
			</select>
		</span>
		
		<span class="itemBlock">
			/ ���� <select name="selPoint" class="select">
				<option value="">��ü</option>
				<option value="1">��</option>
				<option value="2">�ڡ�</option>
				<option value="3">�ڡڡ�</option>
				<option value="4">�ڡڡڡ�</option>
			</select>
		</span>
		<br />
		<label class="itemBlock">�ı⳻�� <input type="text" name="keyword" size="20" value="<%=searchKeyword%>" class="text" /> </label> /
		<span class="itemBlock">
			�˻��Ⱓ
			<input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			<input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> &nbsp;
		</span>
		<span class="itemBlock">
			<label><input type="checkbox" name="chkTerm" value="Check" <% if chkTerm="Check" then Response.Write "checked" %> onClick="swChkTerm(this)" />�Ⱓ��ü</label>&nbsp;
			<label><input type="checkbox" name="chkFirst" <% if chkFirst="on" then Response.Write "checked" %> />ù�ı�</label>&nbsp;
			<label><input type="checkbox" name="photomode" <% IF blnPhotomode="on" Then response.write "checked" %> />�����ǰ�ı�</label>
		</span>
		<script type="text/javascript">
			document.frm.selStatus.value="<%=selStatus%>";
			document.frm.selPoint.value="<%=selPoint%>";
			var CAL_Start = new Calendar({
				inputField : "sDt", trigger    : "sDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDt", trigger    : "eDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</table>
</form>
<br>
<form name="frm_list" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="selStatus" value="<%=selStatus%>">
<input type="hidden" name="SearchKey1" value="<%=SearchKey1%>">
<input type="hidden" name="SearchKey2" value="<%=SearchKey2%>">
<input type="hidden" name="sDt" value="<%=sDt%>">
<input type="hidden" name="eDt" value="<%=eDt%>">
<input type="hidden" name="chkTerm" value="<%=chkTerm%>">
<input type="hidden" name="srtMethod" value="<%=srtMethod%>">
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="left">
		<% if selStatus<>"Y" then %><input type="button" value="���ú���" onClick="doSubmit('restore')" class="button" style="margin:3px 10px 0 0;" /><% end if %>
		<% if selStatus<>"N" then %><input type="button" value="���û���" onClick="doSubmit('delete')" class="button" style="margin:3px 10px 0 0;" /><% end if %>
	</td>
	<td align="right">
	<%
		if Not(oGoodUsing.FAvgTotalPoint="" or isnull(oGoodUsing.FAvgTotalPoint)) then
			Response.Write "<font color=darkred><b>�������</b></font>: "
			Response.Write "<b>���� " & formatNumber(oGoodUsing.FAvgTotalPoint,2) & "��</b> / "
			Response.Write "��� " & formatNumber(oGoodUsing.FAvgFunctionPoint,2) & "�� / "
			Response.Write "������ " & formatNumber(oGoodUsing.FAvgDesignPoint,2) & "�� / "
			Response.Write "���� " & formatNumber(oGoodUsing.FAvgPricePoint,2) & "�� / "
			Response.Write "������ " & formatNumber(oGoodUsing.FAvgSatisfyPoint,2) & "��"
			Response.Write "<span style=""margin:0 10px;"">|</span>"
		end if

		'// ���� �ٿ�ε�
		dim imax, imin, exSize
		exSize = 20000
		if oGoodUsing.FTotalCount>0 then
	%>
		<select name="selDCnt" id="selDCnt" class="select" style="height:20px;vertical-align:top;">
			<%for lp =1 To Int(oGoodUsing.FTotalCount/exSize)+1
					imin = ((lp-1)*exSize)+1
					if lp <  Int(oGoodUsing.FTotalCount/exSize)+1 then
					imax = lp*exSize
					else
					imax = oGoodUsing.FTotalCount
					end if
			%>
			<option value="<%=lp%>"><%=imin%>~<%=imax%></option>
			<%Next%>
		</select>
		<input type="button" class="button_s" value="�����ٿ�ε�" onclick="itemgoodusingExcelDown();">
	<%	end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ���� ��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		�˻���� : <b><%=FormatNumber(oGoodUsing.FTotalCount,0)%></b>
		&nbsp;
		������ : <b><%= page %>/<%=FormatNumber(oGoodUsing.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="tt_sel" onclick="chgSel_on_off()"></td>
	<td>ī�װ�</td>
	<td><%
		if srtMethod="iidDcd" then
			Response.Write "<a href=javascript:ChangeSort('iidAcd') title='��ǰ�ڵ� �ø�����'>��ǰ�ڵ��</a>"
		elseif srtMethod="iidAcd" then
			Response.Write "<a href=javascript:ChangeSort('iidDcd') title='��ǰ�ڵ� ��������'>��ǰ�ڵ��</a>"
		else
			Response.Write "<a href=javascript:ChangeSort('iidDcd') title='��ǰ�ڵ� ��������'>��ǰ�ڵ�</a>"
		end if
	%></td>
	<td>��ǰ��</td>
	<td>��ǰ�ɼ�</td>
	<td><%
		if srtMethod="sprcDcd" then
			Response.Write "<a href=javascript:ChangeSort('sprcAcd') title='�ǸŰ� �ø�����'>�ǸŰ���</a>"
		elseif srtMethod="sprcAcd" then
			Response.Write "<a href=javascript:ChangeSort('sprcDcd') title='�ǸŰ� ��������'>�ǸŰ���</a>"
		else
			Response.Write "<a href=javascript:ChangeSort('sprcDcd') title='�ǸŰ� ��������'>�ǸŰ�</a>"
		end if
	%></td>
	<td>�ۼ���</td>
	<td width="44"><%
		if srtMethod="pntDcd" then
			Response.Write "<a href=javascript:ChangeSort('pntAcd') title='���� �ø�����'>������</a>"
		elseif srtMethod="pntAcd" then
			Response.Write "<a href=javascript:ChangeSort('pntDcd') title='���� ��������'>������</a>"
		else
			Response.Write "<a href=javascript:ChangeSort('pntDcd') title='���� ��������'>����</a>"
		end if
	%></td>
	<td>�ı⳻��</td>
	<td width="80"><%
		if srtMethod="idxDcd" then
			Response.Write "<a href=javascript:ChangeSort('idxAcd') title='�ۼ��� �ø�����'>�ۼ��ϡ�</a>"
		elseif srtMethod="idxAcd" then
			Response.Write "<a href=javascript:ChangeSort('idxDcd') title='�ۼ��� ��������'>�ۼ��ϡ�</a>"
		else
			Response.Write "<a href=javascript:ChangeSort('idxDcd') title='�ۼ��� ��������'>�ۼ���</a>"
		end if
	%></td>
	<td width="40">����</td>
</tr>
<%
	if oGoodUsing.FResultCount=0 then
%>
<tr>
	<td colspan="11" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ��ǰ�ıⰡ �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oGoodUsing.FResultCount - 1
			if oGoodUsing.FitemList(lp).FisUsing="Y" then
				strDel = "<font color=darkblue>�Ϲ�</font>"
			else
				strDel = "<font color=darkred>����</font>"
			end if
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="lineSel" value="<%=oGoodUsing.FitemList(lp).FUsingId%>"></td>
	<td>
		<%=oGoodUsing.FitemList(lp).FCateName%>
	</td>
	<td>
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=oGoodUsing.FitemList(lp).Fitemid%>" target="_blank" title="��ǰ���� ����">
		<%=oGoodUsing.FitemList(lp).Fitemid%></a>
	</td>
	<td>
		<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=oGoodUsing.FitemList(lp).Fitemid%>" target="_blank" title="��ǰ���� ����">
		<%=db2html(oGoodUsing.FitemList(lp).Fitemname)%></a>
	</td>
	<td>
		<% 
			if oGoodUsing.FitemList(lp).Fitemoption="0000" then
				Response.Write "<font color=""gray"">����</font>"
			else
		%>
		[<%=oGoodUsing.FitemList(lp).Fitemoption%>]
		<%=chkIIF(oGoodUsing.FitemList(lp).Fitemoptionname<>"",db2html(oGoodUsing.FitemList(lp).Fitemoptionname),"<font color=""lightgray"">��ǥ�þ���</font>")%>
		<% end if %>
	</td>
	<td>
		<%= formatNumber(oGoodUsing.FItemList(lp).FSellPrice, 0) %>
	</td>
	<td>
		<%= printUserId(oGoodUsing.FItemList(lp).Fuserid, 2, "*") %>
	</td>
	<td align="left" onmouseover="fnPointView(this,'on')" onmouseout="fnPointView(this,'off')" style="cursor:help;">
		<div style="position:absolute;border:2px solid #DDD; background-color:#FFF; width:120px; display:none; padding:5px;">
			<b>����</b>: <% for lp2=1 to oGoodUsing.FitemList(lp).FtotalPoint: Response.Write "<img src=http://fiximage.10x10.co.kr/web2008/category/icon_star.gif>":next%><br />
			���: <% for lp2=1 to oGoodUsing.FitemList(lp).FPointFunction: Response.Write "<img src=http://fiximage.10x10.co.kr/web2008/category/icon_star.gif>":next%><br />
			������: <% for lp2=1 to oGoodUsing.FitemList(lp).FPointDesign: Response.Write "<img src=http://fiximage.10x10.co.kr/web2008/category/icon_star.gif>":next%><br />
			����: <% for lp2=1 to oGoodUsing.FitemList(lp).FPointPrice: Response.Write "<img src=http://fiximage.10x10.co.kr/web2008/category/icon_star.gif>":next%><br />
			������: <% for lp2=1 to oGoodUsing.FitemList(lp).FPointSatisfy: Response.Write "<img src=http://fiximage.10x10.co.kr/web2008/category/icon_star.gif>":next%>
		</div>
		<% for lp2=1 to oGoodUsing.FitemList(lp).FtotalPoint: Response.Write "<img src=http://fiximage.10x10.co.kr/web2008/category/icon_star.gif>":next%>
	</td>
	<td>
		<%=db2html(oGoodUsing.FitemList(lp).Fcontents)%>
		<% IF oGoodUsing.FitemList(lp).FImageIcon1<>"" Then %>
			<div><img src="<%= oGoodUsing.FitemList(lp).FImageIcon1 %>" border="0" width="50" height="50" onClick="showimage('<%=oGoodUsing.FitemList(lp).FImageIcon1%>');" style="cursor:pointer;"></div>
		<% End IF %>
		<% IF oGoodUsing.FitemList(lp).FImageIcon2<>"" Then %>
			<div><img src="<%= oGoodUsing.FitemList(lp).FImageIcon2 %>" border="0" width="50" height="50" onClick="showimage('<%=oGoodUsing.FitemList(lp).FImageIcon2%>');" style="cursor:pointer;"></div>
		<% End IF %>
	</td>
	<td><%=left(oGoodUsing.FitemList(lp).Fregdate,10)%></td>
	<td><%=strDel%></td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr>
	<td colspan="11" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!-- ������ ���� -->
	<%
		if oGoodUsing.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oGoodUsing.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oGoodUsing.StartScrollPage to oGoodUsing.FScrollCount + oGoodUsing.StartScrollPage - 1

			if lp>oGoodUsing.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>[" & lp & "]</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
			end if

		next

		if oGoodUsing.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>

<%
set oGoodUsing = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->