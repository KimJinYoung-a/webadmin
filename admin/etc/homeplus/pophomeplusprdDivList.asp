<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<%
Dim ohomeplus, i, page, infodiv, CateName, searchName
Dim prdDivAllNm, ismapDFT, prdDivAllCode, cateAllNm, ismapDISP
page		= request("page")
infodiv		= request("infodiv")
CateName	= request("CateName")
searchName	= request("searchName")
ismapDFT	= request("ismapDFT")
ismapDISP	= request("ismapDISP")
If page = ""	Then page = 1

'// ��� ����
Set ohomeplus = new CHomeplus
	ohomeplus.FPageSize 	= 20
	ohomeplus.FCurrPage	= page
	ohomeplus.FInfodiv	= infodiv
	ohomeplus.FCateName	= CateName
	ohomeplus.FsearchName = searchName
	ohomeplus.FRectIsMappingDFT	= ismapDFT
	ohomeplus.FRectIsMappingDISP = ismapDISP
	ohomeplus.getTenHomeplusprdDivList
%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		frm.page.value = pg;
		frm.submit();
	}

	// �˻�
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

	// Homeplus ī�װ� ��Ī �˾�
	function popHomeplusprddivMap(mode,infodiv,cdl,cdm,cds,categbn) {
		var pCM = window.open("pophomeplusPrddivMap.asp?mode="+mode+"&infodiv="+infodiv+"&cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&categbn="+categbn,"popprdDivMap","width=600,height=400,scrollbars=yes,resizable=yes");
		pCM.focus();
	}

	function pop_itemmodi(cdl,cdm,cds,infodiv) {
		var pIM = window.open("/admin/itemmaster/itemlist.asp?cdl="+cdl+"&cdm="+cdm+"&cds="+cds+"&infodivYN=Y&infodiv="+infodiv+"&sellyn=Y","popItemmodi","width=1200,height=500,scrollbars=yes,resizable=yes");
		pIM.focus();
	}
//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		���ظ�Ī���� :
		<select name="ismapDFT" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(ismapDFT="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(ismapDFT="N","selected","")%>>�̸�Ī</option>
		</select>
		���ø�Ī���� :
		<select name="ismapDISP" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(ismapDISP="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(ismapDISP="N","selected","")%>>�̸�Ī</option>
		</select> /
		<select name="infodiv" class="select">
			<option value="" >===��ü====</option>
			<option value="01" <%=chkIIF(infodiv="01","selected","")%>>01.�Ƿ�</option>
			<option value="02" <%=chkIIF(infodiv="02","selected","")%>>02.����/�Ź�</option>
			<option value="03" <%=chkIIF(infodiv="03","selected","")%>>03.����</option>
			<option value="04" <%=chkIIF(infodiv="04","selected","")%>>04.�м���ȭ(����/��Ʈ/�׼�����)</option>
			<option value="05" <%=chkIIF(infodiv="05","selected","")%>>05.ħ����/Ŀư</option>
			<option value="06" <%=chkIIF(infodiv="06","selected","")%>>06.����(ħ��/����/��ũ��/DIY��ǰ)</option>
			<option value="07" <%=chkIIF(infodiv="07","selected","")%>>07.������(TV��)</option>
			<option value="08" <%=chkIIF(infodiv="08","selected","")%>>08.������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)</option>
			<option value="09" <%=chkIIF(infodiv="09","selected","")%>>09.��������(������/��ǳ��)</option>
			<option value="10" <%=chkIIF(infodiv="10","selected","")%>>10.�繫����(��ǻ��/��Ʈ��/������)</option>
			<option value="11" <%=chkIIF(infodiv="11","selected","")%>>11.���б��(������ī�޶�/ķ�ڴ�)</option>
			<option value="12" <%=chkIIF(infodiv="12","selected","")%>>12.��������(MP3/���ڻ��� ��)</option>
			<option value="13" <%=chkIIF(infodiv="13","selected","")%>>13.�޴���</option>
			<option value="14" <%=chkIIF(infodiv="14","selected","")%>>14.������̼�</option>
			<option value="15" <%=chkIIF(infodiv="15","selected","")%>>15.�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)</option>
			<option value="16" <%=chkIIF(infodiv="16","selected","")%>>16.�Ƿ���</option>
			<option value="17" <%=chkIIF(infodiv="17","selected","")%>>17.�ֹ��ǰ</option>
			<option value="18" <%=chkIIF(infodiv="18","selected","")%>>18.ȭ��ǰ</option>
			<option value="19" <%=chkIIF(infodiv="19","selected","")%>>19.�ͱݼ�/����/�ð��</option>
			<option value="20" <%=chkIIF(infodiv="20","selected","")%>>20.��ǰ(����깰)</option>
			<option value="21" <%=chkIIF(infodiv="21","selected","")%>>21.������ǰ</option>
			<option value="22" <%=chkIIF(infodiv="22","selected","")%>>22.�ǰ���ɽ�ǰ</option>
			<option value="23" <%=chkIIF(infodiv="23","selected","")%>>23.�����ƿ�ǰ</option>
			<option value="24" <%=chkIIF(infodiv="24","selected","")%>>24.�Ǳ�</option>
			<option value="25" <%=chkIIF(infodiv="25","selected","")%>>25.��������ǰ</option>
			<option value="26" <%=chkIIF(infodiv="26","selected","")%>>26.����</option>
			<option value="27" <%=chkIIF(infodiv="27","selected","")%>>27.ȣ��/��� ����</option>
			<option value="28" <%=chkIIF(infodiv="28","selected","")%>>28.������Ű��</option>
			<option value="29" <%=chkIIF(infodiv="29","selected","")%>>29.�װ���</option>
			<option value="30" <%=chkIIF(infodiv="30","selected","")%>>30.�ڵ��� �뿩 ����(����ī)</option>
			<option value="31" <%=chkIIF(infodiv="31","selected","")%>>31.��ǰ�뿩 ����(������, ��, ����û���� ��)</option>
			<option value="32" <%=chkIIF(infodiv="32","selected","")%>>32.��ǰ�뿩 ����(����, ���ƿ�ǰ, ����ǰ ��)</option>
			<option value="33" <%=chkIIF(infodiv="33","selected","")%>>33.������ ������(����, ����, ���ͳݰ��� ��)</option>
			<option value="34" <%=chkIIF(infodiv="34","selected","")%>>34.��ǰ��/����</option>
			<option value="35" <%=chkIIF(infodiv="35","selected","")%>>35.��Ÿ</option>
		</select>&nbsp;&nbsp;
		<select name="CateName" class="select">
			<option>=��ü=</option>
			<option value="cdlnm" <%=chkIIF(CateName="cdlnm","selected","")%>>��з���</option>
			<option value="cdmnm" <%=chkIIF(CateName="cdmnm","selected","")%>>�ߺз���</option>
			<option value="cdsnm" <%=chkIIF(CateName="cdsnm","selected","")%>>�Һз���</option>
		</select>
		<input type="text" name="searchName" size="20" value="<%=searchName%>">
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="�˻�" onclick="serchItem()">
	</td>
</tr>
</table>
</form>
<p>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>Homeplus ī�װ� ����</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<p>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=ohomeplus.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="5">�ٹ����� ��ǰ����������� �з� ī�װ�</td>
	<td colspan="2">Homeplus ����ī�װ�</td>
	<td colspan="2">Homeplus ����ī�װ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>���<br>��ǰ��</td>
	<td>�ڵ�</td>
	<td>Homeplus ����ī�װ�(�ѱ�)</td>
	<td>�ڵ�</td>
	<td>Homeplus ����ī�װ�(�ѱ�)</td>
</tr>
<% If ohomeplus.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to ohomeplus.FresultCount - 1
			prdDivAllNm = ohomeplus.FItemList(i).FhDiv_Name & ">" & ohomeplus.FItemList(i).FhGROUP_Name & ">" & ohomeplus.FItemList(i).FhDEPT_Name & ">" & ohomeplus.FItemList(i).FhCLASS_Name & ">" & ohomeplus.FItemList(i).FhSUB_NAME
			prdDivAllCode = "<font color='RED'>"&ohomeplus.FItemList(i).FhDIVISION&"</font><font color='ORANGE'>"&ohomeplus.FItemList(i).FhGROUP&"</font>"&ohomeplus.FItemList(i).FhDEPT&"<font color='GREEN'>"&ohomeplus.FItemList(i).FhCLASS&"</font><font color='BLUE'>"&ohomeplus.FItemList(i).FhSUBCLASS&"</font>"
			cateAllNm = ohomeplus.FItemList(i).Fdepth2Nm & ">" & ohomeplus.FItemList(i).Fdepth3Nm & ">" &  ohomeplus.FItemList(i).Fdepth4Nm & ">" &  ohomeplus.FItemList(i).Fdepth5Nm & ">" &  ohomeplus.FItemList(i).Fdepth6Nm
%>
<tr align="center" height="25" bgcolor="#FFFFFF">
	<td><%= ohomeplus.FItemList(i).Finfodiv %></td>
	<td><%= ohomeplus.FItemList(i).FtenCDLName %></td>
	<td><%= ohomeplus.FItemList(i).FtenCDMName %></td>
	<td><%= ohomeplus.FItemList(i).FtenCDSName %></td>
	<td onclick="javascript:pop_itemmodi('<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','<%= ohomeplus.FItemList(i).Finfodiv %>');" style="cursor:pointer;"><%= ohomeplus.FItemList(i).Ficnt %></td>
	<% If ohomeplus.FItemList(i).FhDIVISION="" OR isNull(ohomeplus.FItemList(i).FhDIVISION) Then %>
	<td colspan="2"><input type="button" class="button" value="Homeplus ����ī�װ� ��Ī" onClick="popHomeplusprddivMap('I','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','dft')"></td>
	<% Else %>
	<td title="<%=prdDivAllNm%>" onClick="popHomeplusprddivMap('U','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','dft')" style="cursor:pointer"><%= prdDivAllCode %></td>
	<td><%=prdDivAllNm%></td>
	<% End If %>

	<% If ohomeplus.FItemList(i).FdepthCode="" OR isNull(ohomeplus.FItemList(i).FdepthCode) Then %>
	<td colspan="2"><input type="button" class="button" value="Homeplus ����ī�װ� ��Ī" onClick="popHomeplusprddivMap('I','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','disp')"></td>
	<% Else %>
	<td title="<%=cateAllNm%>" onClick="popHomeplusprddivMap('U','<%= ohomeplus.FItemList(i).Finfodiv %>','<%= ohomeplus.FItemList(i).FtenCateLarge %>','<%= ohomeplus.FItemList(i).FtenCateMid %>','<%= ohomeplus.FItemList(i).FtenCateSmall %>','disp')" style="cursor:pointer"><%= ohomeplus.FItemList(i).FdepthCode %></td>
	<td><%=cateAllNm%></td>
	<% End If %>
</tr>
<%
		Next
	End If
%>
</table>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If ohomeplus.HasPreScroll Then %>
		<a href="javascript:goPage('<%= ohomeplus.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + ohomeplus.StartScrollPage to ohomeplus.FScrollCount + ohomeplus.StartScrollPage - 1 %>
			<% If i > ohomeplus.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If ohomeplus.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<% Set ohomeplus = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->