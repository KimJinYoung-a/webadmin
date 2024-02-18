<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : main_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2008.04.11 ������ : �Ǽ������� ����
'			2009.04.19 �ѿ�� 2009�� �°� ����
'           2009.12.21 ������ : ���ں� �÷��� ���� ��� �߰�
'			2012.02.08 ������ : �̴ϴ޷� ��ü
'           2013.09.28 ������ : 2013������ - �߰����� �ʵ� �߰�
'           2015.04.07 ������ : 2015������ - �߰����� �ʵ� �߰�
'           2018-01-15 ����ȭ : ���� PC��� ���� �߰�
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_enjoyContentsManageCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim isusing, fixtype, validdate, prevDate
Dim idx, poscode, reload, gubun, edid
Dim culturecode
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	gubun = request("gubun")

	isusing = request("isusing")
	fixtype = request("fixtype")
	validdate= request("validdate")
	prevDate = request("prevDate")

	culturecode = request("eC")

	if idx="" then idx=0

	Response.write culturecode

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainEnjoyContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneEnjoyMainContents

	If gubun = "" Then
		gubun = "index"
	End If

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

	function SaveMainContents(frm){
	    if (frm.BGColor.value==""){
	        alert('������ ���� ���� �ϼ���.');
	        frm.BGColor.focus();
	        return;
	    }

	    if (frm.Evt_Code.value==""){
	        alert('�̺�Ʈ ��ȣ�� �Է� �ϼ���.');
	        frm.Evt_Code.focus();
	        return;
	    }

		if (frm.Evt_Title.value==""){
	        alert('�̺�Ʈ ����ī�Ǹ� �Է� �ϼ���.');
	        frm.Evt_Title.focus();
	        return;
	    }
		
		if(!maxLengthCheck("Evt_Title","����ī��",48))
		{
			frm.Evt_Title.focus();
			return;
		}

		if (frm.Evt_Subcopy.value==""){
	        alert('�̺�Ʈ ����ī�Ǹ� �Է� �ϼ���.');
	        frm.Evt_Subcopy.focus();
	        return;
	    }

		if(!maxLengthCheck("Evt_Subcopy","����ī��",80))
		{
			frm.Evt_Title.focus();
			return;
		}

		if (frm.Item1.value==""){
	        alert('��ǰ 1���� �Է� �ϼ���.');
	        frm.Item1.focus();
	        return;
	    }

		if (frm.Item2.value==""){
	        alert('��ǰ 2���� �Է� �ϼ���.');
	        frm.Item2.focus();
	        return;
	    }

		if (frm.Item3.value==""){
	        alert('��ǰ 3���� �Է� �ϼ���.');
	        frm.Item3.focus();
	        return;
	    }

	    if (frm.startdate.value.length!=10){
	        alert('�������� �Է�  �ϼ���.');
	        return;
	    }

	    if (frm.enddate.value.length!=10){
	        alert('�������� �Է�  �ϼ���.');
	        return;
	    }

	    if (frm.startdate.value>frm.enddate.value){
	        alert('�������� �����Ϻ��� ������ �ȵ˴ϴ�.');
	        return;
	    }

	    if (confirm('���� �Ͻðڽ��ϱ�?')){
	        frm.submit();
	    }
	}

	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}

	//function getOnLoad(){
	//    ChangeLinktype(frmcontents.linktype.value);
	//}

	//window.onload = getOnLoad;

	function ChangeGubun(comp){
	    location.href = "?gubun=<%=gubun%>&poscode=" + comp.value;
	    // nothing;
	}


	function ChangeGroupGubun(comp){
	    location.href = "?gubun=" + comp.value;
	    // nothing;
	}

	function cultureloadpop(){
		winLast = window.open('pop_culturelist.asp','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}

	//�����ڵ� ����
	function selColorChip(bg,cd) {
		var i;
		document.frmcontents.BGColor.value= bg;
		for(i=1;i<=11;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}

	//-- jsLastEvent : ���� �̺�Ʈ �ҷ����� --//
	function jsLastEvent(){
	  winLast = window.open('pop_event_lastlist.asp','pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}

	function fnviewimg(num){
		var itemid = $("#Item"+num).val();
		$.ajax({
			type: "POST",
			url: "/admin/sitemaster/lib/item_image_view_act.asp",
			data: "Itemid="+itemid,
			cache: false,
			success: function(message) {
				$("#img"+num).empty().html("<img src='"+message+"' width='100' height='100' border='0'>");
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	/**
	 * ����Ʈ ���� �Է°��� ���ڼ� üũ
	 * 
	 * @param id : tag id 
	 * @param title : tag title
	 * @param maxLength : �ִ� �Է°��� �� (byte)
	 * @returns {Boolean}
	 */
	function maxLengthCheck(id, title, maxLength){
		 var obj = $("#"+id);
		 if(maxLength == null) {
			 maxLength = obj.attr("maxLength") != null ? obj.attr("maxLength") : 1000;
		 }
		 
		 if(Number(byteCheck(obj)) > Number(maxLength)) {
			 alert(title + "��(��) �Է°��ɹ��ڼ��� �ʰ��Ͽ����ϴ�.\n(����, ����, �Ϲ� Ư������ : " + maxLength + " / �ѱ�, ����, ��Ÿ Ư������ : " + parseInt(maxLength/2, 10) + ").");
			 obj.focus();
			 return false;
		 } else {
			 return true;
		}
	}

	/**
	 * ����Ʈ�� ��ȯ  
	 * 
	 * @param el : tag jquery object
	 * @returns {Number}
	 */
	function byteCheck(el){
		var codeByte = 0;
		for (var idx = 0; idx < el.val().length; idx++) {
			var oneChar = escape(el.val().charAt(idx));
			if ( oneChar.length == 1 ) {
				codeByte ++;
			} else if (oneChar.indexOf("%u") != -1) {
				codeByte += 2;
			} else if (oneChar.indexOf("%") != -1) {
				codeByte ++;
			}
		}
		return codeByte;
	}

</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doMainEnjoyContentsReg.asp" onsubmit="return false;">
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>

        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">����</td>
    <td>
		<table id="colorselect1">
			<tr>
				<td><table id='cline11' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#FBD65C" Or oMainContents.FOneItem.FBGColor="" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#FBD65C' style="font-size:8px"><a href='javascript:selColorChip("#FBD65C",11)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline1' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#FFB137" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#FFB137' style="font-size:8px"><a href='javascript:selColorChip("#FFB137",1)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline2' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#DCBBEC" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#DCBBEC' style="font-size:8px"><a href='javascript:selColorChip("#DCBBEC",2)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline3' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#C1BEFE" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#C1BEFE' style="font-size:8px"><a href='javascript:selColorChip("#C1BEFE",3)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline4' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#B9DAFA" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#B9DAFA' style="font-size:8px"><a href='javascript:selColorChip("#B9DAFA",4)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline5' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#AAE9DB" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#AAE9DB' style="font-size:8px"><a href='javascript:selColorChip("#AAE9DB",5)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline6' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#CBF09C" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#CBF09C' style="font-size:8px"><a href='javascript:selColorChip("#CBF09C",6)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline7' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#DFDFDF" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#DFDFDF' style="font-size:8px"><a href='javascript:selColorChip("#DFDFDF",7)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td><table id='cline8' border='0' cellpadding='0' cellspacing='1' bgcolor='<% If oMainContents.FOneItem.FBGColor="#C0C0C0" Then %>#DD3300<% Else %>#dddddd<% End If %>'><tr><td bgcolor='#FFFFFF'><table border='0' cellpadding='0' cellspacing='2'><tr><td bgcolor='#C0C0C0' style="font-size:8px"><a href='javascript:selColorChip("#C0C0C0",8)' onfocus='this.blur()'><img src='http://testwebadmin.10x10.co.kr/images/space.gif' alt='��Ȳ' width='12' height='12' border='0'></a></td></tr></table></td></tr></table></td>
				<td>�����Է�<input type="text" name="BGColor" value="<%=oMainContents.FOneItem.FBGColor%>"></td>
			</tr>
		</table>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">Ÿ��</td>
    <td>
    	<input type="radio" name="Evt_Type" value="1"<% If oMainContents.FOneItem.FEvt_Type="1" Then Response.write " checked" %> onClick="jsLastEvent()"> �̺�Ʈ �ҷ�����
		<input type="radio" name="Evt_Type" value="2"<% If oMainContents.FOneItem.FEvt_Type="2" Then Response.write " checked" %>> �����Է�
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">�̺�Ʈ</td>
    <td>
    	�̺�Ʈ �ڵ� : <input type="text" name="Evt_Code" value="<%=oMainContents.FOneItem.FEvt_Code%>"> <a href="">�̸�����</a>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">����ī��</td>
    <td>
		<input type="text" name="Evt_Title" id="Evt_Title" value="<%=oMainContents.FOneItem.FEvt_Title%>" size="50"><br>
		������ : <input type="text" name="Evt_Discount" value="<%=oMainContents.FOneItem.FEvt_Discount%>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">����ī��</td>
    <td>
    	<input type="text" name="Evt_Subcopy" id="Evt_Subcopy" value="<%=oMainContents.FOneItem.FEvt_Subcopy%>" size="80">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��ǰ1</td>
  <td>
  	<input type="text" name="Item1" id="Item1" value="<%=oMainContents.FOneItem.FItem1%>"> <a href="javascript:fnviewimg(1);">�̸�����</a><br>
	<div id="img1"><img src="<% = GetItemImageLoad(oMainContents.FOneItem.FItem1) %>" width="100" height="100" border="0"></div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��ǰ2</td>
  <td>
  	<input type="text" name="Item2" id="Item2" value="<%=oMainContents.FOneItem.FItem2%>"> <a href="javascript:fnviewimg(2);">�̸�����</a><br>
	<div id="img2"><img src="<% = GetItemImageLoad(oMainContents.FOneItem.FItem2) %>" width="100" height="100" border="0"></div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��ǰ3</td>
  <td>
  	<input type="text" name="Item3" id="Item3" value="<%=oMainContents.FOneItem.FItem3%>"> <a href="javascript:fnviewimg(3);">�̸�����</a><br>
	<div id="img3"><img src="<% = GetItemImageLoad(oMainContents.FOneItem.FItem3) %>" width="100" height="100" border="0"></div>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">������</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FStartDate)%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	<script type="text/javascript">
	var CAL_Start = new Calendar({
		inputField : "startdate",
		trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		},
		bottomBar: true,
		dateFormat: "%Y-%m-%d"
	});
	</script>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">������</td>
  <td>
  	<input type="text" name="EndDate" id="enddate" value="<%=chkiif(idx=0,prevDate,oMainContents.FOneItem.FEndDate)%>">
	<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	<script type="text/javascript">
	var CAL_End = new Calendar({
		inputField : "enddate",
		trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		},
		bottomBar: true,
		dateFormat: "%Y-%m-%d"
	});
	</script>
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�켱����</td>
  <td>
  	<input type="text" name="DispOrder" value="<%=oMainContents.FOneItem.FDispOrder%>">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">��뿩��</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oMainContents.FOneItem.FIsusing="Y" Or oMainContents.FOneItem.FIsusing="" Then Response.write " checked" %>> �����
	<input type="radio" name="Isusing" value="N"<% If oMainContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> ������
  </td>
</tr>
<% If oMainContents.FOneItem.FRegUser<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="150" bgcolor="#DDDDFF">�۾���</td>
  <td>
  	�۾��� : <%=oMainContents.FOneItem.FRegUser %><br>
	�����۾��� : <%=oMainContents.FOneItem.FLastUser %>
  </td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<%
set oMainContents = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
