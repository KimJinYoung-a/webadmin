<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : ��ǰ�� ��� �귣�� ���� ��� ������
'	History		: 2017.01.20 ���¿� ����
'				  2017.02.27 �ѿ�� ����(����� �ΰ� �ϳ� ����, ������ Ÿ�� ��ũ��Ʈ ����)
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/board/brand_noticeCls.asp"-->

<%
Dim i, mode, menupos, sDt, sTm, eDt, eTm, sdate, edate, gubun, notice_title, notice_text, makerid, brandid
Dim srcSDT , srcEDT, stdt, eddt, idx, isusing, regdate, SearchGubun
Dim sqlstr, sqlsearch, arrlist, resultcount, infiniteregyn, opart
	menupos				=	requestcheckvar(request("menupos"),10)
	idx				=	requestcheckvar(request("idx"),32)
	mode			=	requestcheckvar(request("mode"),4)
	srcSDT			=	requestcheckvar(request("sDt"),10)
	srcEDT			=	requestcheckvar(request("eDt"),10)
'	sdate			=	requestcheckvar(request("sdate"),10)
'	edate			=	requestcheckvar(request("edate"),10)
	gubun			=	requestcheckvar(request("gubun"),1)
	makerid			=	requestcheckvar(request("makerid"),32)
	brandid			=	requestcheckvar(request("brandid"),32)
	isusing			=	requestcheckvar(request("isusing"),1)
	regdate			=	requestcheckvar(request("regdate"),32)
	notice_text		=	requestcheckvar(request("notice_text"),256)
	infiniteregyn	=	requestcheckvar(request("infiniteregyn"),1)

	set opart = new CBrandNotice
		opart.fnGetBrandNoticeList

if idx = "" then 
	mode="NEW"
else
	mode="EDIT"
end if

if mode="EDIT" then
	if idx <> "" then
		sqlsearch = sqlsearch & " and idx="& idx &""
	end if
		
		sqlstr = "select top 1"
		sqlstr = sqlstr & " idx, sdate, edate, isusing, regdate, gubun, makerid, brandid, notice_title , notice_text, infiniteregyn"
		sqlstr = sqlstr & " from db_board.dbo.tbl_brand_notice_list with (nolock)"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by idx desc"

		'response.write sqlstr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
		resultcount = rsget.recordcount
		
	if not rsget.EOF then
		arrlist = rsget.getrows()
	end if
	
	rsget.close

	idx				= arrlist(0,0)
	sdate 			= arrlist(1,0)
	edate 			= arrlist(2,0)
	isusing 		= arrlist(3,0)
	regdate 		= arrlist(4,0)
	gubun 			= arrlist(5,0)
	makerid 		= arrlist(6,0)
	brandid			= arrlist(7,0)
	notice_title	= arrlist(8,0)
	notice_text		= arrlist(9,0)
	infiniteregyn	= arrlist(10,0)
end if

if Not(sdate="" or isNull(sdate)) then
	sDt = left(sdate,10)
	sTm = Num2Str(hour(sdate),2,"0","R") &":"& Num2Str(minute(sdate),2,"0","R")
else
	if srcSDT<>"" then
		sDt = left(srcSDT,10)
	else
		sDt = date
	end if
	sTm = "00:00"
end if

if Not(edate="" or isNull(edate)) then
	eDt = left(edate,10)
	eTm = Num2Str(hour(edate),2,"0","R") &":"& Num2Str(minute(edate),2,"0","R")
else
	if srcEDT<>"" then
		eDt = left(srcEDT,10)
	else
		eDt = date
	end if
	eTm = "23:59"
end If

%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	
function frmedit(){
	if(frm.brandid.value==""){
		alert("�귣��ID�� �����ϴ�.");
		frm.brandid.focus();
		return;
	}

	if(frm.SearchGubun.value==""){
		alert("���������� ������ �ּ���");
		frm.SearchGubun.focus();
		return;
	}

	if(frm.notice_title.value==""){
		alert("���������� �Է��� �ּ���.");
		frm.notice_title.focus();
		return;
	}

	if(frm.notice_text.value==""){
		alert("���������� �Է��� �ּ���.");
		frm.notice_text.focus();
		return;
	}


	if(frm.StartDate.value==""){
		alert("���� �������� ������ �ּ���");
		frm.StartDate.focus();
		return;
	}

	if(frm.EndDate.value==""){
		alert("���� �������� ������ �ּ���");
		frm.EndDate.focus();
		return;
	}

// �̰� ���� ����. ������ ȣȯ���� ���� ����.	//2017.02.27 �ѿ��
//	var filter = ['<P', '<p', '<IMG', '<img', 'class=', 'style=', 'src=', 'SRC='];
//	var matchcnt = 0;
//	var txt = $('#notice_text').val();
//	for( var i in filter ){
//		try{
//			var compare = txt.match( filter[i] );
//			console.log( compare.index );
//			alert( '����� ���� �Ǿ��ֽ��ϴ�. - ' + filter[i] );
//			matchcnt++;
//			if( matchcnt > 0 ) return;
//			} catch( err ) {
//			console.log( '���' );
//		}
//	}

	alert('���������� ��¥�� ��ġ�� ���� �ֱ� ����� ���������� ��� �˴ϴ�.');
	frm.submit();
}

$(function(){
	//�޷´�ȭâ ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
	$("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showCurrentAtPos: 1,
		showOn: "button",
		<% if Idx<>"" then %>maxDate: "<%=eDt%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
		}
	});
	$("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showOn: "button",
		<% if Idx<>"" then %>minDate: "<%=sDt%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

function chghicprogbn(comp){
    var frm=comp.form;
	location.href="/admin/board/brand_notice/brand_notice_write.asp?idx=<%= idx %>&gubun="+comp;
}

</script>

<img src="/images/icon_arrow_link.gif"> <b>�귣�� ���� ���</b>
<form name="frm" method="post" action="brand_notice_proc.asp" style="margin:0px;">
<input type = "hidden" name = "idx" value = "<%=idx %>">
<input type = "hidden" name = "mode" value = "<%=mode %>">
<input type="hidden" name="menupos" value="<%=menupos %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if mode = "EDIT"  then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȣ</td>
		<td colspan="2"><%=idx%></td>
	</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�귣��ID</td>
		<td colspan="2">
			<%	drawSelectBoxDesignerWithName "brandid", brandid %>&nbsp;&nbsp;&nbsp;<% if brandid <> "" then %><a href="http://www.10x10.co.kr/street/street_brand.asp?makerid=<%= brandid %>" target="blank" >�귣�� ������ ></a><% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
		<td colspan="2">
			<select name="SearchGubun" ><%''onChange='chghicprogbn(this.value)'%>
				<option value ="" style="color:blue">�� ��</option>
				<option value="1" <% If "1" = cstr(gubun) Then %> selected <% End if %>>�Ϲݰ���</option>
				<option value="2" <% If "2" = cstr(gubun) Then %> selected <% End if %>>��۰���</option>
				<option value="3" <% If "3" = cstr(gubun) Then %> selected <% End if %>>��Ÿ����</option>
			</select>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">���� ����</td>
		<td colspan="2">
			<input type="text" size="62" name="notice_title" value="<%=notice_title%>" id="notice_title" onBlur="checkFieldColor(this);" onFocus="clearFieldColor(this);"/>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">���� ����</td>
		<td colspan="2">
			<textarea title="���� ����" cols="63" rows="5" name="notice_text" id="notice_text" ><%=notice_text%></textarea>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">���� �Ⱓ</td>
		<td colspan="2">
			<% if mode = "NEW" then %>
				<input type="text" id="sDt" name="StartDate" size="10" value="<%=stdt%>" />
				<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
				<input type="text" id="eDt" name="EndDate" size="10" value="<%=eddt%>" />
				<input type="text" name="eTm" size="8" value="<%=eTm%>" />
			<% else %>
				<input type="text" id="sDt" name="StartDate" size="10" value="<%=sDt%>" />
				<input type="text" name="sTm" size="8" value="<%=sTm%>" /> ~
				<input type="text" id="eDt" name="EndDate" size="10" value="<%=eDt%>" />
				<input type="text" name="eTm" size="8" value="<%=eTm%>" />
			<% end if %>
			<Br><input type="checkbox" name="infiniteregyn" id="infiniteregyn" <% if infiniteregyn = "Y" then%>checked<% end if %>>���� ���� (��ó���)
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
		<td colspan="2">
			<input type="radio" name="isusing" value="Y" <%=chkiif(isusing = "Y","checked","")%> checked />����� &nbsp;&nbsp;&nbsp; 
			<input type="radio" name="isusing" value="N"  <%=chkiif(isusing = "N","checked","")%>/>������
		</td>
	</tr>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="3">
		<%' if gubun <> "" then %>
			<% if mode = "EDIT" or mode = "NEW" then %>
				<input type="button" class="button" uname="editsave" value="����" onclick="frmedit()" />	
			<% end if %>
				<input type="button" class="button" name="editclose" value="���" onclick="self.close()" />
		<%' end if %>
		</td>
	</tr>
</table>
</form>
<% set opart = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
