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
<!-- #include virtual="/lib/classes/sitemasterclass/gnb_main_ContentsManageCls.asp" -->
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
		set oMainContents = new CGNBContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneEventMainContents
	
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

		if(!maxLengthCheck("Evt_Title","�̺�Ʈ ����ī��",40))
		{
			frm.Evt_Title.focus();
			return;
		}

		if (frm.Evt_Subcopy.value==""){
	        alert('�̺�Ʈ ����ī�Ǹ� �Է� �ϼ���.');
	        frm.Evt_Subcopy.focus();
	        return;
	    }

		if(!maxLengthCheck("Evt_Subcopy","�̺�Ʈ ����ī��",80))
		{
			frm.Evt_Subcopy.focus();
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
	function jsLastEvent(num){
	  winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
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
<form name="frmcontents" method="post" action="doMainEventReg.asp" onsubmit="return false;">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">Idx</td>
    <td>
        <% if oMainContents.FOneItem.Fidx<>"" then %>
        <%= oMainContents.FOneItem.Fidx %>
        <input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
        <% else %>
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">�̺�Ʈ �ڵ�</td>
	<td><input type="text" name="Evt_Code" value="<%=oMainContents.FOneItem.FEvt_Code%>"> <a href="javascript:jsLastEvent(1);">�ҷ�����</a></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">����ī��</td>
	<td><input type="text" name="Evt_Title" id="Evt_Title" value="<%=oMainContents.FOneItem.FEvt_Title%>" size="50" maxlength="30"></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">������</td>
	<td><input type="text" name="Evt_Discount" value="<%=oMainContents.FOneItem.FEvt_Discount%>"></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">���� ������</td>
	<td><input type="text" name="Evt_Coupon" value="<%=oMainContents.FOneItem.FEvt_Coupon%>"></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="#DDDDFF">����ī��</td>
	<td><input type="text" name="Evt_Subcopy" id="Evt_Subcopy" value="<%=oMainContents.FOneItem.FEvt_Subcopy%>" size="70" maxlength="40"></td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">������</td>
  <td>
  	<input type="text" name="StartDate" id="startdate" value="<%=oMainContents.FOneItem.FStartDate%>">
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
  <td width="100" bgcolor="#DDDDFF">������</td>
  <td>
  	<input type="text" name="EndDate" id="enddate" value="<%=oMainContents.FOneItem.FEndDate%>">
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
  <td width="100" bgcolor="#DDDDFF">�켱����</td>
  <td>
  	<input type="text" name="DispOrder" value="<%=oMainContents.FOneItem.FDispOrder%>">
  </td>
</tr>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">��뿩��</td>
  <td>
  	<input type="radio" name="Isusing" value="Y"<% If oMainContents.FOneItem.FIsusing="Y" Or oMainContents.FOneItem.FIsusing="" Then Response.write " checked" %>> �����
	<input type="radio" name="Isusing" value="N"<% If oMainContents.FOneItem.FIsusing="N" Then Response.write " checked" %>> ������
  </td>
</tr>
<% If oMainContents.FOneItem.FRegUser<>"" Then %>
<tr bgcolor="#FFFFFF">
  <td width="100" bgcolor="#DDDDFF">�۾���</td>
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