<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_reviewCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterClass/main_TopReviewCls.asp"-->

<%
dim SearchKey1,SearchKey2, makerid, cdl, cdm, cds, sDt, eDt, chkTerm, blnPhotomode, keyword, itemname
dim sellyn
SearchKey1 = Request("SearchKey1")
SearchKey2 = request("SearchKey2")
makerid     = requestCheckvar(request("makerid"),32)
cdl = Request("cdl")
cdm = Request("cdm")
cds = Request("cds")
sDt = Request("sDt")
eDt = Request("eDt")
chkTerm = Request("chkTerm")
blnPhotomode = Request("photomode")

sellyn  = request("sellyn")
keyword = request("keyword")
itemname = request("itemname")


Dim page, idx
idx = Request("idx")
page = Request("page")
If page="" Then page = 1


	if sDt="" and chkTerm="" then sDt = DateAdd("d",-1,date())
	if eDt="" and chkTerm="" then eDt = date()
dim oeventuserlist , i

	set oeventuserlist = new Ceventuserlist
		oeventuserlist.FPagesize = 20
		oeventuserlist.FCurrPage = page
		oeventuserlist.fSearchKey1 = SearchKey1
		oeventuserlist.fSearchKey2 = SearchKey2
		oeventuserlist.FRectMakerid	= makerid
		oeventuserlist.FRectCDL	=	cdl
		oeventuserlist.FRectCDM	=	cdm
		oeventuserlist.FRectCDS	=	cds
		oeventuserlist.FRectStartDt = sDt
		oeventuserlist.FRectEndDt = eDt
		oeventuserlist.FRectPhotoMode = blnPhotomode
		oeventuserlist.FRectSellYN       = sellyn
		oeventuserlist.FRectKeyword		= keyword

		'oeventuserlist.FRectItemName	= itemname
		oeventuserlist.Feventuserlist99()


	dim omainreview
	Set omainreview = new CSearchKeyWord
	omainreview.FRectidx = idx

	if idx<>"" then
		omainreview.GetSearchreview
	end if



function lcate(aa)
	dim l_cate
	l_cate = aa
	select case l_cate
	case "010"
		response.write "�����ι���"
	case "020"
		response.write "���ǽ�/���μ�ǰ"
	case "025"
		response.write "������"
	case "030"
		response.write "Ű��Ʈ"
	case "035"
		response.write "����/���"
	case "040"
		response.write "����"
	case "045"
		response.write "����/��Ȱ"
	case "050"
		response.write "Ȩ/����"
	case "055"
		response.write "�к긯"
	case "060"
		response.write "Űģ"
	case "070"
		response.write "����/����/���"
	case "075"
		response.write "��Ƽ"
	case "080"
		response.write "Women"
	case "090"
		response.write "Men"
	case "100"
		response.write "���̺�"
	case "110"
		response.write "����ä��"
	end select

end function

%>

<script language='javascript'>
	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="main_review_write.asp";
		document.frm.submit();
	}
	function choice(uid,cmt,iid,Lcate,Mcate, iname)
	{
		document.frm1.userid.value= uid;
		document.frm1.comment.value=cmt;
		document.frm1.itemid.value=iid;
		document.frm1.cate_large.value=Lcate;
		document.frm1.cate_mid.value=Mcate;
		document.frm1.itemname.value=iname;
	}
	function goSubmit()
	{
		// id �Է¿��� �˻�
		if(!document.frm1.userid.value) {
			alert("���� Ű���带 �Է����ּ���.");
			document.frm1.userid.focus();
			return;
		}
		// �ڸ�Ʈ �Է¿��� �˻�
		if(!document.frm1.comment.value) {
			alert("Ű���� Ŭ���� �̵��� ��ũ�� �Է����ּ���.");
			document.frm1.comment.focus();
			return;
		}

		// ���� �Է¿��� �˻�
		if(!document.frm1.sortNo.value) {
			alert("ǥ�� ������ �Է����ּ���.\n�� ������ �����̸� �������� ������ �����ϴ�.");
			document.frm1.sortNo.focus();
			return;
		}

		<% if idx="" then %>
		if(confirm("�ۼ��Ͻ� ������ ����Ͻðڽ��ϱ�?")) {
			document.frm1.mode.value="add";
			document.frm1.action="doMainReview.asp";
			document.frm1.submit();
		}
		<% else %>
		if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?")) {
			document.frm1.mode.value="modify";
			document.frm1.action="doMainReview.asp";
			document.frm1.submit();
		}
		<% end if %>
	}



	// ���� ���� ����
	function chgStatus(v)
	{
		document.frm.selStatus.value=v;
		document.frm.submit();
	}

	// ��ǰ�� �˾�
	function viewItemInfo(iid)
	{
		var PpUp = window.open("<%=wwwurl%>/common/PopZoomItem.asp?itemid="+ iid +"&pop=pop","itemInfo","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,width=720,height=444");
		PpUp.focus();
	}

	// ���Ĺ�� ����
	function ChangeSort(smtd)	{
		document.frm.srtMethod.value=smtd;
		document.frm.submit();
	}

	// ��ü ����,���
	function chgSel_on_off()
	{
		var frm = document.frm_list;
		if (frm.lineSel.length)
		{
			for(var i=0;i<frm.lineSel.length;i++)
			{
				frm.lineSel[i].checked=frm.tt_sel.checked;
			}
		}
		else
		{
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

	//�̹��� ����
	function showimage(img){
		var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
	}

// ī�װ� ����� ���
function changecontent(){
}


</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!-- �˻� ���� -->
<!-- <form name="searchfrm" method="post" >
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td width="100" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
		<td align="left">
			&nbsp;ItemID: <input type="text" name="seachbox" value="" size="10">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.searchfrm.submit();">
		</td>
	</tr>
</table>
</form> -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		<!-- #include virtual="/common/module/categoryselectbox.asp"--> &nbsp; /&nbsp; �Ǹ�:<% oeventuserlist.drawSelectBoxSell "sellyn", sellyn %><br>
		���̵� <input type="text" name="SearchKey1" size="12" value="<%=SearchKey1%>" class="text">
		/ ��ǰ��ȣ <input type="text" name="SearchKey2" size="12" value="<%=SearchKey2%>" class="text">
		/ �귣��ID <%	drawSelectBoxDesignerWithName "makerid", makerid %>
		<br>
		������ �˻� : <input type="text" class="text" name="keyword" value="<%=keyword%>" size="40"><font color="gray" size="2">(����:�������ֽ��ϴ�.)</font>
		<br>
		�˻��Ⱓ
		<input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<input type="checkbox" name="chkTerm" value="Check" <% if chkTerm="Check" then Response.Write "checked" %> onClick="swChkTerm(this)">�Ⱓ��ü
		<input type="checkbox" name="photomode" <% IF blnPhotomode="on" Then response.write "checked" %>>�����ǰ�ı�
		<script language="javascript">
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
</form>
</table>
<!-- �˻� �� -->

<!-- ����Ʈ ���� -->
<form name="frm1" method="post" action="doMainReview.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cate_large" value="">
<input type="hidden" name="cate_mid" value="">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="5" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if idx="" then %>
		<font color="red"><b>�ڸ�Ʈ ���</b></font>
		<% else %>
		<font color="red"><b>�ڸ�Ʈ ����</b></font>
		<% end if%>
	</td>
</tr>
<% if idx<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�Ϸù�ȣ</td>
	<td align="left" colspan="3"><input type="text" name="idx" value="<%=idx%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" colspan="2">User ID</td>
	<td align="left" colspan="3"><input type="text" name="userid" value="<% if idx<> "" then Response.Write omainreview.FitemList(0).fuserid %>" size="18" readonly maxlength="18" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" colspan="2">item ID</td>
	<td align="left" colspan="3"><input type="text" name="itemid" value="<% if idx<> "" then Response.Write omainreview.FitemList(0).fitemid %>" size="18" readonly maxlength="18" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" colspan="2">item Name</td>
	<td align="left" colspan="3"><input type="text" name="itemname" value="<% if idx<> "" then Response.Write omainreview.FitemList(0).fitemname %>" size="50" readonly maxlength="50" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" colspan="2">�ڸ�Ʈ</td>
	<td align="left" colspan="3">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="5"><input type="text" bgcolor="#707080" name="comment" value="<% if idx<>"" then Response.Write omainreview.FitemList(0).fcomment%>" size="200" readonly class="text"></td>
		<tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>" colspan="2">ǥ�ü���</td>
	<td align="left" colspan="3"><input type="text" name="sortNo" value="<% if idx<>"" then Response.Write omainreview.FitemList(0).FsortNo: else Response.Write "99" %>" size="3" class="text"></td></td>
</tr>
	<% if oeventuserlist.ftotalcount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17">
			�˻���� : <b><%= oeventuserlist.FTotalCount %></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" class="button" value="����" onClick="goSubmit()">
					<input type="button" class="button" value="���" onClick="self.history.back()">
		</td>

	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">�ۼ���</td>
		<td align="center" width="100">ī�װ�</td>
		<td align="center">��ǰ��</td>
		<td align="center" width="1200">Comment</td>
		<td align="center"width="100">�ۼ���</td>
    </tr>

	<% for i=0 to oeventuserlist.FResultCount-1 %>
    	<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><a href="javascript:choice('<%= oeventuserlist.flist(i).fuserid %>','<%= chrbyte(oeventuserlist.flist(i).fcontents,300,"Y") %>','<%= oeventuserlist.flist(i).fitemid %>','<%= oeventuserlist.flist(i).fcate_large %>','<%= oeventuserlist.flist(i).fcate_mid %>','<%= oeventuserlist.flist(i).fitemname %>')"><%= oeventuserlist.flist(i).fuserid %></a></td>
			<td><% lcate(oeventuserlist.flist(i).fcate_large)%></td>
			<td align="center"><a href="<%= wwwurl %>/shopping/category_prd.asp?itemid=<%=oeventuserlist.flist(i).fitemid  %>" target="_blank">[<%= oeventuserlist.flist(i).fitemid %>] <%= oeventuserlist.flist(i).fitemname %></a></td>
			<td align="left" style="padding:10px"><a href="javascript:choice('<%= oeventuserlist.flist(i).fuserid %>','<%= chrbyte(oeventuserlist.flist(i).fcontents,300,"Y") %>','<%= oeventuserlist.flist(i).fitemid %>','<%= oeventuserlist.flist(i).fcate_large %>','<%= oeventuserlist.flist(i).fcate_mid %>','<%= oeventuserlist.flist(i).fitemname %>')"><%= oeventuserlist.flist(i).fcontents %></a>
		<% IF oeventuserlist.flist(i).FImageIcon1<>"" Then %>
			<br><img src="<%= oeventuserlist.flist(i).FImageIcon1 %>" border="0" width="50" height="50" onClick="showimage('<%=oeventuserlist.flist(i).FImageIcon1%>');" style="cursor:pointer;">&nbsp;&nbsp;
		<% End IF %>
		<% IF oeventuserlist.flist(i).FImageIcon2<>"" Then %>
			<img src="<%= oeventuserlist.flist(i).FImageIcon2 %>" border="0" width="50" height="50" onClick="showimage('<%=oeventuserlist.flist(i).FImageIcon2%>');" style="cursor:pointer;">
		<% End IF %>
				</td>
			<td><%= left(oeventuserlist.flist(i).fregdate,10) %></td>


    	</tr>
	<% next %>



</table>

</form>

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom" bgcolor="FFFFFF">
			<td align="center">
			<!-- ������ ���� -->
			<%
				if oeventuserlist.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oeventuserlist.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + oeventuserlist.StartScrollPage to oeventuserlist.FScrollCount + oeventuserlist.StartScrollPage - 1

					if i>oeventuserlist.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oeventuserlist.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->