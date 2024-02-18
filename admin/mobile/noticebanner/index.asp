<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : index.asp
' Discription : ����� ����Ʈ �˸����
' History : 2013.04.01 ����ȭ ����
'			2016.07.21 �ѿ�� ����
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_noticebanner.asp" -->
<%	
Dim userlevel , isusing, sdate, edate, page, i, vParam
	page = request("page")
	sdate = RequestCheckVar(request("startday"),13)
	edate = RequestCheckVar(request("endday"),13)
	userlevel = RequestCheckVar(request("userlevel"),13)
	isusing = RequestCheckVar(request("isusing"),13)

If sdate = "" Then sdate = date
If edate = "" Then edate = DateAdd("d" , 7 , date) '�⺻ ��������
if page="" then page=1

vParam = "&sdate="&sdate&"&edate="&edate&"&userlevel="&userlevel&"&isusing="&isusing

dim oNoticebanner
set oNoticebanner = new CMainbanner
	oNoticebanner.FPageSize		= 20
	oNoticebanner.FCurrPage		= page
	oNoticebanner.FSearchSdate = sdate
	oNoticebanner.FSearchEdate = edate
	oNoticebanner.Fisusing			= isusing
	oNoticebanner.Fuserlevel		= userlevel
	oNoticebanner.GetContentsList()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>

// ���� ���� ���� 7��
function chkdate(v){
	var frm = document.frm;
	var nowdate = new Date();
	var year  = nowdate.getFullYear();
	var month = nowdate.getMonth() + 1; // 1��=0,12��=11�̹Ƿ� 1 ����
	var day   = nowdate.getDate();

	if (("" + month).length == 1) { month = "0" + month; }
	if (("" + day).length   == 1) { day   = "0" + day;   }

	today  = year + "-" + month + "-" + day; //������

	if (v == "N"){
		frm.startday.value = today;
		frm.endday.value =  today;
	}else if (v =="T"){
		frm.startday.value = today;
		frm.endday.value = mathdate(today,"1");
	}else if (v =="W"){
		frm.startday.value = mathdate(today,"-7");
		frm.endday.value = today;
	}
}

// ��¥ ���
function mathdate(date,v){
		var input1 = date;
		var input2 = v;
 		var dateinfo = input1.split("-");
		var src = new Date(dateinfo[0], dateinfo[1]-1, dateinfo[2]);

		src.setDate(src.getDate() + parseInt(input2));
		var year = src.getFullYear();
	    var month = src.getMonth() + 1;
		var date = src.getDate();

		if(month<10) month = "0" + month;
 
		if(date<10) date = "0" + date;
 
		var result = year + "-" + month + "-" + date;

		return result;
}

//����
function jsmodify(v){
	location.href = "nb_insert.asp?menupos=<%=menupos%>&idx="+v;
}

function jschgusing(v,idx){
	location.href = "nb_proc.asp?iidx="+idx+"&isusing="+v+"&mode=chg";
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
      	maxDate: "<%=edate%>",
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
      	minDate: "<%=sdate%>",
    	onClose: function( selectedDate ) {
    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
    	}
    });
});

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<div style="padding-bottom:10px;">
			* ��뿩�� :&nbsp;&nbsp;<% DrawSelectBoxUsingYN "isusing",isusing %>&nbsp;&nbsp;&nbsp;
			* ���뱸�� :&nbsp;&nbsp;<% DrawselectboxUserLevel "userlevel", userlevel, "" %>
			</div>
			<div>
	       	* ��ȸ�Ⱓ :&nbsp;&nbsp; 
			<input type="button"  class="button_s" value="����" onclick="chkdate('N');">&nbsp;
			<input type="button"  class="button_s" value="����" onclick="chkdate('T');">&nbsp;
			<input type="button"  class="button_s" value="���� 7��" onclick="chkdate('W');">&nbsp;

			<input type="text" id="sDt" name="startday" size="10" value="<%=sdate%>" /> ~
			<input type="text" id="eDt" name="endday" size="10" value="<%=edate%>" />
			</div>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
		</td>
	</tr>
</form>	
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="nb_insert.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�� ��ϼ� : <b><%=oNoticebanner.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oNoticebanner.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="7%">��ȣ(idx)</td>
    <td width="7%">�����</td>
    <td width="16%">����</td>
    <td width="17%">����Ⱓ</td>
    <td width="7%">�켱����</td>
    <td width="10%">��뿩��</td>
    <td >��ޱ���</td>
    <td width="10%">�����</td>
</tr>
<% 
	Dim tempSdate , tempStime ,  tempEdate , tempEtime

	for i=0 to oNoticebanner.FResultCount-1 

		''��¥ �ð� �и�
		tempSdate = ""
		tempEdate = ""
		tempStime = ""
		tempEtime = ""
		If Len(oNoticebanner.FItemList(i).Fstartday) <= 10 Or Len(oNoticebanner.FItemList(i).Fendday) <= 10  Then
			tempSdate = oNoticebanner.FItemList(i).Fstartday
			tempEdate = oNoticebanner.FItemList(i).Fendday
		Else
			tempSdate = Left(oNoticebanner.FItemList(i).Fstartday,10)
			tempStime = Trim(right(oNoticebanner.FItemList(i).Fstartday,11))
			tempEdate = Left(oNoticebanner.FItemList(i).Fendday,10)
			tempEtime = Trim(right(oNoticebanner.FItemList(i).Fendday,11))
		End If 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(oNoticebanner.FItemList(i).Fisusing="1","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=oNoticebanner.FItemList(i).Fiidx%>');" style="cursor:pointer;"><%=oNoticebanner.FItemList(i).Fiidx%></td>
    <td><%=oNoticebanner.FItemList(i).Fwriter%></td>
    <td><%=oNoticebanner.FItemList(i).Ftitle%></td>
    <td><%=tempSdate%> ~ <%=tempEdate%><% If tempStime <> "" Or tempEtime <>"" Then %><br/>(<%=tempStime%> ~ <%=tempEtime%>)<% End If %></td>
    <td><%=oNoticebanner.FItemList(i).Fsorting%></td>
    <td><%=chkiif(oNoticebanner.FItemList(i).Fisusing="0","������","�����")%>&nbsp;<input type="button" value="����"  class="button_s" onclick="jschgusing('<%=oNoticebanner.FItemList(i).Fisusing%>','<%=oNoticebanner.FItemList(i).Fiidx%>');"/></td>
    <td><%=oNoticebanner.FItemList(i).FutnArr%></td>
    <td><%=Left(oNoticebanner.FItemList(i).Fwritedate,10)%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin-top:20px;padding-right:80px;" border="0">
	<tr>
		<td align="center" width="60%">
			<% if oNoticebanner.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oNoticebanner.StartScrollPage-1 %><%=vParam%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oNoticebanner.StartScrollPage to oNoticebanner.StartScrollPage + oNoticebanner.FScrollCount - 1 %>
				<% if (i > oNoticebanner.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oNoticebanner.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %><%=vParam%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oNoticebanner.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %><%=vParam%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set oNoticebanner = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->