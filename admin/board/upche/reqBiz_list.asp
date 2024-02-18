<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
'###########################################################
' Description : ��ü ������޼�
' History : 2008.09.01 �ѿ�� ����/�߰�
'			2014.05.13 ������ ����
'###########################################################
%>

<%
dim i, j,ix
dim page,gubun, onlymifinish
dim research, searchkey,catevalue, dispCate,maxDepth
dim ipjumYN , catemid ,catelarge, sellgubun
Dim workid
Dim iid
	page 			= requestCheckvar(request("pg"),10)
	gubun 			= requestCheckvar(request("gubun"),2)
	onlymifinish 	= requestCheckvar(request("onlymifinish"),3)
	research 		= requestCheckvar(request("research"),3)
	searchkey 		= requestCheckvar(request("searchkey"),32)
	catevalue		= requestCheckvar(request("catevalue"),3)
	ipjumYN			= requestCheckvar(request("ipjumYN"),1)
	catemid 		= requestCheckvar(request("catemidbox"),3)
	catelarge 		= requestCheckvar(request("catelargebox"),3)
	dispCate		= requestCheckVar(Request("disp"),16) 
	maxDepth		= 2
	sellgubun			= requestCheckvar(request("sellgubun"),1)
	workid			= requestCheckvar(request("workid"),34)
	iid             = requestCheckVar(Request("iid"),9) 
	
 
	 gubun="02"
	if research="" and onlymifinish="" then onlymifinish="on"		
	if (page = "") then page = "1"
 	 

dim companyrequest
set companyrequest = New CCompanyRequest
	companyrequest.PageSize = 20
	companyrequest.CurrPage = CInt(page)
	companyrequest.ScrollCount = 10
	companyrequest.FReqcd=gubun
	companyrequest.FOnlyNotFinish = onlymifinish
	companyrequest.FRectSearchKey = searchkey
	companyrequest.FRectCatevalue = catevalue
	companyrequest.FipjumYN = ipjumYN
	companyrequest.FRectDispCate = dispCate
	companyrequest.FRectSellgubun = sellgubun
	companyrequest.FRectWorkid = workid
	companyrequest.FRectID=iid
	companyrequest.list

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function checkComp(comp){
 
		document.location.href='/admin/board/upche/req_list.asp?menupos=<%=menupos%>&gubun=02&disp=&catevalue=';
	 
}

//����Ʈ
function printpage(id){
	
	var printpage;
	printpage = window.open("/admin/board/upche/req_print.asp?id=" +id, "printpage","width=1024,height=768,scrollbars=yes,resizable=yes");
	printpage.focus();

}

function delitem(id){
	
	if (confirm("�����Ͻðڽ��ϱ�?.") ==true)
		frmdel.mode.value="del";
		frmdel.id.value=id;
		frmdel.submit();
}
function MovePage(page){
	frm.pg.value=page;
	frm.research.value="<%=research %>";
	//frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="/admin/board/upche/req_list.asp";
	frm.submit();
}

function ViewPage(id){
 
		var winView = window.open("/admin/board/upche/req_view2.asp?id="+id,"popReq","width=1024,height=768,scrollbars=yes,resizable=yes");
	 
	winView.focus();
/*
	var winView = window.open("about:blank;","popReq","width=1024,height=768,scrollbars=yes,resizable=yes");
	frm.id.value=id;
	frm.pg.value=<%=page%>;
	frm.research.value="<%=research %>";
	//frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.target = "popReq";
	 
		frm.action="/admin/board/upche/req_view2.asp";
	 
	frm.submit();
*/
}

function DownPage(id,sFN){
	  var winFD = window.open("<%=uploadImgUrl%>/linkweb/company/downcorequest.asp?idx="+id+"&sFN="+sFN,"popFD","");
    winFD.focus();
} 

function changecontent() {
	frm.pg.value="1";
	frm.submit();
}

</script> 
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="id" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="pg" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
	 
		<input type="hidden" name="catevalue" value="">
		�������� : 
		<select name="sellgubun" class="select">
			<option value="">��ü</option>
			<option value="1" <%= Chkiif(sellgubun="1", "selected", "") %> >��������</option>
			<option value="2" <%= Chkiif(sellgubun="2", "selected", "") %> >����������</option>
			<option value="3" <%= Chkiif(sellgubun="3", "selected", "") %> >���������� �� ���θ�� ����</option>
			<option value="4" <%= Chkiif(sellgubun="4", "selected", "") %> >��ȭ�̺�Ʈ ����</option>
			<option value="5" <%= Chkiif(sellgubun="5", "selected", "") %> >��� �� �ַ�� ���� ����</option>
			<option value="6" <%= Chkiif(sellgubun="6", "selected", "") %> >������</option>
		</select>&nbsp;&nbsp;
		����� : 
		<% DrawWorkIdCombo "workid", workid %>
	 
		<select name="ipjumYN" class="a">
			<option value="">�Ϸᱸ��</option>
			<option value="Y" <% if ipjumYN="Y" then response.write "selected" %>>�����Ϸ�</option>
			<option value="N" <% if ipjumYN="N" then response.write "selected" %>>�̿Ϸ�</option>
		</select>
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="changecontent();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		 
		&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="onlymifinish" <% if onlymifinish="on" then response.write "checked" %> >ó���ȵȸ��
		&nbsp;&nbsp;&nbsp;&nbsp;
		��ü�� <input type="text" name="searchkey" value="<%= searchkey %>">	
		
		&nbsp;&nbsp;&nbsp;&nbsp;
		�۹�ȣ <input type="text" name="iid" value="<%= iid %>" size=6>		
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">	
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if companyrequest.resultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= companyrequest.TotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= companyrequest.TotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">��ȣ</td>
    <td align="center">��û��</td>
    <td align="center">����</td>
    <td align="center">ó����</td>
    <td align="center">��������</td> 
    <td align="center">ȸ��URL</td>
    <td align="center">�亯����</td>
    <td align="center">���</td>
    </tr>
	<% for i = 0 to (companyrequest.ResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%=companyrequest.results(i).id%></td>
	    <td align="center" nowrap><%= FormatDate(companyrequest.results(i).regdate, "0000-00-00") %></td>
	    <td>[<%= companyrequest.code2name(companyrequest.results(i).reqcd) %>] <%= companyrequest.results(i).companyname %></td>
	    <td align="center" nowrap>
	        <% if (IsNull(companyrequest.results(i).finishdate) = true) then %>
	      <font color="red">�̿Ϸ�</font>
	        <% else %>
	      <%= FormatDate(companyrequest.results(i).finishdate, "0000-00-00") %>
	        <% end if %>
	    </td>
	    <td align="center">
	    	<%if companyrequest.results(i).ipjumYN="Y" then response.write "�����Ϸ�" %>
	    	<%if companyrequest.results(i).ipjumYN="N" then response.write "N" %>
	    	</td>
	  	 
	  	<td align="center">
	  		<a href="<%IF left(companyrequest.results(i).companyurl,4)<>"http" then%>http://<%END IF%><%= companyrequest.results(i).companyurl%>" target="_blank"><%= companyrequest.results(i).companyurl%></a>
	  	</td> 
	  	<td align="center">
	  		<% if companyrequest.commentcheck(companyrequest.results(i).replycomment)="Y" then %>
	  		Y
	  		<% else %>
	  		<font color="red">N</font>
	  		<% end if %>
	  	</td>
	  	<td align="center" nowrap>
		  	<input type="button" value="����" class="button" onclick="javascript:ViewPage(<%= companyrequest.results(i).id %>);">
		   
		  	<%if companyrequest.results(i).attachfile <> "" then%><input type="button" value="÷�����ϴٿ�" class="button" onclick="javascript:DownPage(<%= companyrequest.results(i).id %>,'<%=companyrequest.results(i).attachfile%>');"><%end if%>
	  	</td>
    </tr>   
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		<% if companyrequest.HasPreScroll then %>
			<a href="javascript:MovePage(<%= companyrequest.StartScrollPage-1 %>);">[prev]</a>
		<% else %>
			[prev]
		<% end if %>

		<% for ix=0 + companyrequest.StartScrollPage to companyrequest.ScrollCount + companyrequest.StartScrollPage - 1 %>
			<% if ix>companyrequest.Totalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
					<font color="red">[<%= ix%>]</font>
				<% else %>
					<a href="javascript:MovePage(<%=ix%>);">[<%= ix %>]</a>
				<% end if %>
		<% next %>

		<% if companyrequest.HasNextScroll then %>
			<a href="javascript:MovePage(<%=ix%>);">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>

<form name="frmdel" method="get" action="cscenter_req_board_act.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="page" value="<%=page%>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->