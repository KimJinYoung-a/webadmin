<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : nb_insert.asp
' Discription : ����� ����Ʈ �˸����
' History : 2013.04.01 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_noticebanner.asp" -->
<%
Dim sdate , edate , idx , mode
Dim writer
Dim temputarr
Dim tempSdate , tempEdate , tempShour , tempEhour

	idx = request("idx")
	writer  = session("ssBctCname")

If idx = "" Then 
	mode = "add" 
	idx = 0
Else 
	mode = "modify" 
End If 

dim oNoticeBannerOne
set oNoticeBannerOne = new CMainbanner
oNoticeBannerOne.FRectIdx = idx
oNoticeBannerOne.GetOneContents()

	Function checksel(val) ''����Ʈ�ڽ� üũ
		temputarr = Split(oNoticeBannerOne.FOneItem.FutArr,",")
		Dim ii  
			For ii = 0 To UBound(temputarr)
				If val = temputarr(ii) Then
					checksel = "checked"
				Exit for
				End If 
			next
	End Function

	''//��¥ �ð� ����
	sdate = oNoticeBannerOne.FOneItem.Fstartday
	edate = oNoticeBannerOne.FOneItem.Fendday

	If Not(sdate="" or isNull(sdate)) then
		tempSdate = Left(sdate,10)
		tempShour = Num2Str(hour(sdate),2,"0","R") & ":" & Num2Str(minute(sdate),2,"0","R") & ":" & Num2Str(second(sdate),2,"0","R")
	else
		tempSdate = date
		tempShour = "00:00:00"
	end if

	If Not(edate="" or isNull(edate)) then
		tempEdate = Left(edate,10)
		tempEhour = Num2Str(hour(edate),2,"0","R") & ":" & Num2Str(minute(edate),2,"0","R") & ":" & Num2Str(second(edate),2,"0","R")
	else
		tempEdate = date
		tempEhour = "23:59:59"
	end if
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript'>
	function jsSubmit(){
		var frm = document.frm;
		if (frm.utArr.value ==""){
			checkval();
		}

		if (frm.utArr.value ==""){
			alert('��� ������ ���� ���� �ϼ���.');
			return;
		}
		
		if (frm.startday.value.length!=10){
			alert('����Ⱓ �������� �Է� �ϼ���.');
			frm.startday.focus();
			return;
		}
		
		if (frm.endday.value.length!=10){
			alert('����Ⱓ �������� �Է� �ϼ���.');
			frm.endday.focus();
			return;
		}

		if (frm.title.value == "" ){
			alert('������ �Է� �ϼ���');
			frm.title.focus();
			return;
		}

		if (frm.sorting.value == "" ){
			alert('�켱������ �Է� �ϼ���');
			frm.sorting.focus();
			return;
		}

		if (frm.text.value == "" ){
			alert('�ؽ�Ʈ ������ �Է� �ϼ���');
			frm.text.focus();
			return;
		}

		if (confirm('���� �Ͻðڽ��ϱ�?')){
			//frm.target = "_blank";
			frm.action = "/admin/mobile/noticebanner/nb_proc.asp";
			frm.submit();
		}
	}

	//���� ���ڼ� ����
	function textCounter(field, countfield, maxlimit) {
		if (field.value.length > maxlimit){ 
			field.blur();
			field.value = field.value.substring(0, maxlimit);
			//alert("15�� �̳����� �����ּ���");
			field.focus();
		}
		else {
			countfield.value = maxlimit - field.value.length;
			document.getElementById("counttxt").innerHTML = maxlimit - countfield.value;
		}
	}
	// ��ޱ��� ��ü����
	function allcheck(){
		var adddot = ",";
		var frm = document.frm;
		var length = frm.usertype.length-1;
		frm.utnArr.value  = "";
		frm.utArr.value = "";
		if ( frm.usertype[0].checked == true ){
			for ( i = 1 ; i <= length ; i++ )
			{
				if ( i ==  length)
				{
					adddot = "";
				}
				frm.usertype[i].checked = true;
				frm.utnArr.value = frm.utnArr.value+ $("input[name='usertype']").eq(i).attr("value2") + adddot;
				frm.utArr.value =  frm.utArr.value + $("input[name='usertype']").eq(i).attr("value") + adddot;
			}
		}else{
			for ( i = 1 ; i <= length ; i++ )
			{
				frm.usertype[i].checked = false;
				frm.utnArr.value = "";
				frm.utArr.value = "";
			}
		}
	}
	// ��ޱ��� ���� ����
	function checkval(){
		var adddot = ",";
		var frm = document.frm;
		var length = frm.usertype.length-1;
		frm.usertype[0].checked = false;
		frm.utnArr.value  = "";
		frm.utArr.value = "";
		for ( i = 1 ; i <= length ; i++ )
		{
			if ( frm.usertype[i].checked )
			{
				if ( i ==  length)
				{
					adddot = "";
				}
				frm.utnArr.value = frm.utnArr.value+ $("input[name='usertype']").eq(i).attr("value2") + adddot;
				frm.utArr.value =  frm.utArr.value + $("input[name='usertype']").eq(i).attr("value") + adddot;
			}
		}
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
	      	maxDate: "<%=tempEdate%>",
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
	      	minDate: "<%=tempSdate%>",
	    	onClose: function( selectedDate ) {
	    		$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
	    	}
	    });
	});
</script>

<table width="800" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frm" method="post">
<input type="hidden" name="utnArr" value="">
<input type="hidden" name="utArr" value="">
<input type="hidden" name="iidx" value="<%=idx%>">
<input type="hidden" name="mode" value="<%=mode%>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��� ����</td>
	<td>
		<input type="checkbox" name="usertype" value="9" onclick="allcheck();" value2="��ü" <% If mode="modify" then%><%=checksel("9")%><%End If %>/> ��ü
		<input type="checkbox" name="usertype" value="5" onclick="checkval();" value2="������" <% If mode="modify" then%><%=checksel("5")%><%End If %>/> ������
		<input type="checkbox" name="usertype" value="0" onclick="checkval();" value2="���ο�" <% If mode="modify" then%><%=checksel("0")%><%End If %>/> ���ο�
		<input type="checkbox" name="usertype" value="1" onclick="checkval();" value2="�׸�" <% If mode="modify" then%><%=checksel("1")%><%End If %>/> �׸�
		<input type="checkbox" name="usertype" value="2" onclick="checkval();" value2="���" <% If mode="modify" then%><%=checksel("2")%><%End If %>/> ���
		<input type="checkbox" name="usertype" value="3" onclick="checkval();" value2="VIP�ǹ�" <% If mode="modify" then%><%=checksel("3")%><%End If %>/> VIP�ǹ�
		<input type="checkbox" name="usertype" value="4" onclick="checkval();" value2="VIP���" <% If mode="modify" then%><%=checksel("4")%><%End If %>/> VIP���
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">���� �Ⱓ</td>
	<td>
		<input type="text" id="sDt" name="startday" size="10" value="<%=tempSdate%>" />
		<input type="text" name="sthh" size="8" value="<%=tempShour%>" /> ~
		<input type="text" id="eDt" name="endday" size="10" value="<%=tempEdate%>" />
		<input type="text" name="edhh" size="8" value="<%=tempEhour%>" />
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">����</td>
	<td><input type="text" class="text"  name="title" size="50" maxlength="40" onKeyDown="textCounter(this.form.title,this.form.remLen,20);" onKeyUp="textCounter(this.form.title,this.form.remLen,20);" value="<%=oNoticeBannerOne.FOneItem.Ftitle%>"/><input type="hidden" name="remLen" value="20"/>&nbsp;<span id="counttxt"/>0</span>�� / 20��</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�켱����</td>
	<td><div style="float:left;"><input type="text" class="text" name="sorting" size="5" maxlength="3" value="<%=oNoticeBannerOne.FOneItem.Fsorting%>"/></div> <div style="float:right;margin-top:5px;margin-right:10px;">�س����� : 99(�ֻ��)~1(���ϴ�)</div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">�ؽ�Ʈ</td>
	<td>
		<table width="80%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFFF" width="15%" align="center">����</td>
				<td><input type="text" class="text" name="text" size="50" maxlength="22" value="<%=oNoticeBannerOne.FOneItem.Ftext%>"/></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2" bgcolor="#87ceeb" align="left">�ִ� 22��</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFFF" width="15%"  align="center">���� ī��</td>
				<td><input type="text" class="text" name="infourl" size="50" value="<%=oNoticeBannerOne.FOneItem.Ftextcopy%>" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2" bgcolor="#87ceeb" align="left">��) Ȯ���Ϸ� ����> ( > �ڵ��Է� ���� )</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td bgcolor="#FFFFFF" width="15%"  align="center">����URL</td>
				<td><input type="text" class="text" name="texturl" size="50" value="<%=oNoticeBannerOne.FOneItem.Ftexturl%>" /></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2" bgcolor="#87ceeb" align="left">��) /event/eventm.asp?eventid=6264 (������ ������� ����α�)</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center">��뿩��</td>
	<td><div style="float:left;"><input type="radio" name="isusing" value="1" <%=chkiif(oNoticeBannerOne.FOneItem.Fisusing = "1","checked","")%>/>����� &nbsp;&nbsp;&nbsp; <input type="radio" name="isusing" value="0" checked <%=chkiif(oNoticeBannerOne.FOneItem.Fisusing = "0","checked","")%>/>������</div> <div style="float:right;margin-top:5px;margin-right:10px;">�� ����� : ���γ��� / ������ : ���γ��� ����</div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="25">�۾���</td>
	<td><div style="float:left"><strong><%=chkiif(mode="add",writer,oNoticeBannerOne.FOneItem.Fwriter)%></strong></div><div style="float:right;margin-right:10px;">�� �۾��ڴ� ������ ���ε� �Ǵ� ������ ����ڰ� ��ϵ˴ϴ�.</div></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFF999" align="center" height="25">���� ������</td>
	<td><div style="float:left"><strong><%=writer%></strong></div><div style="float:right;margin-right:10px;">�� ���� �����ڴ� ������ ������Ʈ ����ڰ� ��ϵ˴ϴ�.</div></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan="2"><input type="button" value=" �� �� " onClick="history.back(-1)"/><input type="button" value=" �� �� " onClick="jsSubmit();"/></td>
</tr>
</form>
</table>
<%
set oNoticeBannerOne = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->