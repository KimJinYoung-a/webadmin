<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/cjmall/cjmallitemcls.asp"-->
<%
Dim ocjmall, i, page, infodiv, CateName, searchName
Dim prdDivAllNm, isMapping
page		= request("page")
infodiv		= request("infodiv")
CateName	= request("CateName")
searchName	= request("searchName")
isMapping	= request("ismap")
If page = ""	Then page = 1

'// ��� ����
Set ocjmall = new CCjmall
	ocjmall.FPageSize 	= 20
	ocjmall.FCurrPage	= page
	ocjmall.Finfodiv	= infodiv
	ocjmall.FCateName	= CateName
	ocjmall.FsearchName	= searchName
	ocjmall.FRectIsMapping	= isMapping
	ocjmall.getTencjmallprdDivList
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

	// cjmall ��ǰ�з� ��Ī �˾�
	function popCjprddivMap(mode,infodiv,cdl,cdm,cds,dno) {
		var pCM = window.open("popcjmallprdDivMap.asp?mode="+mode+"&infodiv="+infodiv+"&cdl="+cdl+"&cdm="+cdm+"&cds="+cds,"popprdDivMap","width=600,height=400,scrollbars=yes,resizable=yes");
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
		��Ī���� :
		<select name="ismap" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(isMapping="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(isMapping="N","selected","")%>>�̸�Ī</option>
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
	<font color="red"><strong>cjmall ��ǰ�з� ����</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=ocjmall.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="5">�ٹ����� ��ǰ����������� �з� ī�װ�</td>
	<td colspan="3">cjmall ��ǰ�з�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>��з�</td>
	<td>�ߺз�</td>
	<td>�Һз�</td>
	<td>���<br>��ǰ��</td>
	<td>�ڵ�</td>
	<td>���з���</td>
	<td>cjmall ��ǰ�з�(�ѱ�)</td>
</tr>
<% If ocjmall.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to ocjmall.FresultCount - 1
			prdDivAllNm = ocjmall.FItemList(i).Fcdl_Name & ">" & ocjmall.FItemList(i).Fcdm_Name & ">" & ocjmall.FItemList(i).Fcds_Name & ">" & ocjmall.FItemList(i).Fcdd_Name
			
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(ocjmall.FItemList(i).FPrdDivIsUsing="Y","#FFFFFF","#CCCCCC") %>">
	<td><%= ocjmall.FItemList(i).Finfodiv %></td>
	<td><%= ocjmall.FItemList(i).FtenCDLName %></td>
	<td><%= ocjmall.FItemList(i).FtenCDMName %></td>
	<td><%= ocjmall.FItemList(i).FtenCDSName %></td>
	<td onclick="javascript:pop_itemmodi('<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>','<%= ocjmall.FItemList(i).Finfodiv %>');" style="cursor:pointer;"><%= ocjmall.FItemList(i).Ficnt %></td>
	<% If ocjmall.FItemList(i).FCddKey="" OR isNull(ocjmall.FItemList(i).FCddKey) Then %>
	<td colspan="3"><input type="button" class="button" value="cjmall ��ǰ�з� ��Ī" onClick="popCjprddivMap('I','<%= ocjmall.FItemList(i).Finfodiv %>','<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>')"></td>
	<% Else %>
	<td title="<%=prdDivAllNm%>" onClick="popCjprddivMap('U','<%= ocjmall.FItemList(i).Finfodiv %>','<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>')" style="cursor:pointer"><%= ocjmall.FItemList(i).FCddKey %></td>
	<td title="<%=prdDivAllNm%>" onClick="popCjprddivMap('U','<%= ocjmall.FItemList(i).Finfodiv %>','<%= ocjmall.FItemList(i).FtenCateLarge %>','<%= ocjmall.FItemList(i).FtenCateMid %>','<%= ocjmall.FItemList(i).FtenCateSmall %>')" style="cursor:pointer"><%= ocjmall.FItemList(i).Fcdd_Name %></td>
	<td><%=prdDivAllNm%></td>
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
		<% If ocjmall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= ocjmall.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + ocjmall.StartScrollPage to ocjmall.FScrollCount + ocjmall.StartScrollPage - 1 %>
			<% If i > ocjmall.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If ocjmall.HasNextScroll Then %>
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
<% Set ocjmall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->