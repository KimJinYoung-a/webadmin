<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->

<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<%
Dim code_large , code_mid , weclass , chgstyle , classtype

	code_large  = RequestCheckvar(request("code_large"),3)
	code_mid   = RequestCheckvar(request("code_mid"),3)
	weclass = RequestCheckvar(request("weclass"),1)

	If weclass = "Y" Then
		chgstyle = "style='display:none;'"
		classtype = "<font color='red'><strong>WeClass ��ü ���� �Է�</storng></font>"
	End If 

'���� ���̵�,���纰 ���� ǥ��
'<option value="������̵�,�����̸�(�Ҽ�),����>���̵�,�����̸�(�Ҽ�)</option>
'''db_academy.dbo.tbl_lec_user ??
public Sub SelectLecturerId()
	dim sqlStr,i
''	sqlStr = "select  c.userid,p.company_name,c.defaultmargine, c.regdate, u.lec_margin, u.mat_margin"
''	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
''	sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
''	sqlStr = sqlStr + "     left join [ACADEMYDB].db_academy.dbo.tbl_lec_user u on c.userid=u.lecturer_id"
''	sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
''	sqlStr = sqlStr + " and c.userdiv < 22" + vbcrlf
''	sqlStr = sqlStr + " and c.userdiv='14'" + vbcrlf

	sqlStr = "select  u.lecturer_id, g.lecturer_name, u.lec_margin, u.mat_margin, u.regdate, u.lecturer_name as brandName"
	sqlStr = sqlStr + " from db_academy.dbo.tbl_lec_User u"
	sqlStr = sqlStr + "     left join db_academy.dbo.tbl_corner_good g"
	sqlStr = sqlStr + "     on u.lecturer_id=g.lecturer_id"
	sqlStr = sqlStr + " where u.lec_yn='Y'" + vbCrlf
	sqlStr = sqlStr + " order by u.lecturer_id"
	
    rsAcademyget.CursorLocation = adUseClient
    rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
	if not rsAcademyget.eof then
			response.write "<select name='temp_lec_id' onchange='javascript:FnLecturerApp(this.value);'>"
			response.write "<option value=''>����</option>"
		for i=0 to rsAcademyget.recordcount-1
			response.write "<option value='" & db2html(rsAcademyget("lecturer_id")) & "," & db2html(rsAcademyget("lecturer_name")) & "," & rsAcademyget("lec_margin") & "," & rsAcademyget("mat_margin") & "," & left(rsAcademyget("regdate"),10) & "'>" & db2html(rsAcademyget("lecturer_id")) & "(" & db2html(rsAcademyget("lecturer_name")) & ") - "&db2html(rsAcademyget("brandName"))&"</option>"
		rsAcademyget.movenext
		next
			response.write "</select>"
	end if
    rsAcademyget.Close
end sub
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--

//submit �⺻���� üũ
function frmsub(frm){

	//var frm=document.lecfrm;

	if (frm.CateCD1.value.length < 1){
		alert('Ŭ���� ������ ������ �ּ���');
		frm.CateCD1.focus();
		return;
	}

//	if (frm.CateCD2.value.length < 1){
//		alert('���� ������ ������ �ּ���');
//		frm.CateCD2.focus();
//		return;
//	}

	if (frm.CateCD3.value.length < 1){
		alert('��� ������ ������ �ּ���');
		frm.CateCD3.focus();
		return;
	}

	if (frm.classlevel.value.length < 1){
		alert('���� ����� ������ �ּ���');
		frm.classlevel.focus();
		return;
	}
	<% if weclass="" or weclass="N" then%>
	if (frm.lec_date.value.length < 1){
		alert('���¿��� �Է��� �ּ���.');
		frm.lec_date.focus();
		return;
	}
	<% end if %>

	if (frm.lec_title.value.length < 1){
		alert('���¸��� �Է��� �ּ���.');
		frm.lec_title.focus();
		return;
	}

	if (frm.lecturer_id.value.length < 1){
		alert('���縦 ������ �ּ���.');
		frm.lecturer_id.focus();
		return;
	}


	//if (frm.lec_cost.value.length < 1|| frm.lec_cost.value==0){
	if (frm.lec_cost.value.length < 1){
		alert('�����Ḧ �Է��� �ּ���.');
		frm.lec_cost.focus();
		return;
	}
	
	//if (frm.buying_cost.value.length < 1|| frm.buying_cost.value==0){
	if (frm.buying_cost.value.length < 1){
		alert('���԰� �ڵ������ ���ּ���.');
		frm.buying_cost.focus();
		return;
	}
//2016/12/13 ������ũ�� ���걸�� ����
    if (frm.code_large.value=="76"){
        if (frm.lecjgubun.value.length<1){
            alert('ī�װ��� ������ũ�� �� ��� �������� �����ϼ���.');
            frm.lecjgubun.focus();
            return;
        }
    }else{
        if (frm.lecjgubun.value=="1"){
            alert('ī�װ��� ������ũ���� �ƴѰ�� �������� �⺻�̳� ����(����) ���� �����ϼ���.');
            frm.lecjgubun.focus();
            return;
        }
    }
	
//������ ���� ���� -_-;

	if (frm.mileage.value.length < 1){
		alert('���ϸ����� �Է��� �ּ���.');
		frm.mileage.focus();
		return;
	}

	//if (frm.mat_contents.value.length < 1){
	//	alert('���� ������ �Է��� �ּ���.');
	//	frm.mat_contents.focus();
	//	return;
	//}

	// ��Ÿ���� �����Ϳ��� ������ �Է�
//	if (sector_1.chk==0){
//		frm.lec_etccontents.value = editor.document.body.innerHTML;
//	} else if(sector_1.chk!=3){
//		frm.lec_etccontents.value = editor.document.body.innerText;
//	}

	frm.submit();
}

//���� ���Եɶ� input box disable ��Ŵ
function Fnmat(){
var frm=document.lecfrm;

	if(frm.matinclude_yn.checked){
		frm.matinclude_yn.value='';
		frm.mat_cost.disabled='on';
	}else{
		frm.matinclude_yn.value='on';
		frm.mat_cost.disabled='';
	}
}

//���纰 ����,���̵�,�Ҽ� ǥ��
function FnLecturerApp(str){
	var varArray;
	varArray = str.split(',');
    
    if (varArray[0]){
    	document.lecfrm.lecturer_id.value = varArray[0];
    	document.lecfrm.lecturer_name.value = varArray[1];
    	document.lecfrm.margin.value = varArray[2];
    	document.lecfrm.mat_margin.value = varArray[3];
    	document.lecfrm.lecturer_regdate.value = varArray[4];
    }else{
        document.lecfrm.lecturer_id.value = "";
        document.lecfrm.lecturer_name.value = "";
    	document.lecfrm.margin.value = 0;
    	document.lecfrm.mat_margin.value = 0;
    	document.lecfrm.lecturer_regdate.value = "";
    }
    
	CalcuAuto(document.lecfrm);
}

//���԰� �ڵ� ��� ǥ��
function CalcuAuto(frm){
	var imargin, isellcash, ibuycash;
	var isellvat, ibuyvat, imileage;

	imargin = frm.margin.value;
	isellcash = frm.lec_cost.value;

	if (imargin.length<1){
		alert('������ �Է��ϼ���.');
		frm.margin.focus();
		return;
	}

	if (isellcash.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.lec_cost.focus();
		return;
	}

	if (!IsDouble(imargin)){
		alert('������ ���ڷ� �Է��ϼ���.');
		frm.margin.focus();
		return;
	}

	if (!IsDigit(isellcash)){
		alert('�ǸŰ��� ���ڷ� �Է��ϼ���.');
		frm.lec_cost.focus();
		return;
	}

	isellvat = 0;
	ibuycash = isellcash - parseInt(isellcash*imargin/100);
	ibuyvat = 0;
	imileage = parseInt((isellcash*1 + frm.mat_cost.value*1)*0.01) ;


	//frm.sellvat.value = isellvat;
	//frm.lec_cost.value = isellvat;
	frm.buying_cost.value=ibuycash;
	//frm.buyvat.value = ibuyvat;
	frm.mileage.value = imileage;
}

function CalcuAutoMaT(frm){

	if (frm.mat_margin.value.length<1){
		alert('���� ������ �Է��ϼ���.');
		frm.mat_margin.focus();
		return;
	}

	if (frm.mat_cost.value.length<1){
		alert('���� �Է��ϼ���.');
		frm.mat_cost.focus();
		return;
	}

	if (!IsDouble(frm.mat_margin.value)){
		alert('������ ���ڷ� �Է��ϼ���.');
		frm.mat_margin.focus();
		return;
	}

	if (!IsDigit(frm.mat_cost.value)){
		alert('����� ���ڷ� �Է��ϼ���.');
		frm.mat_cost.focus();
		return;
	}

	var ibuycash = frm.mat_cost.value*1 - parseInt(frm.mat_cost.value*frm.mat_margin.value/100);

	frm.mat_buying_cost.value=ibuycash;
	frm.mileage.value = parseInt((frm.lec_cost.value*1 + frm.mat_cost.value*1)*0.01) ;
}


//���� ���� �̹��� ��� ���� �����ֱ�.
function showimgyn(){
	var frm = imagetag.style
	frm.display='block';
}

//���½ð� �߰��� input box �߰� 
function addtime(){
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<div>"));
	timetbl.insertAdjacentText("BeforeEnd","�����Ͻ� ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Sday' value=''>"));
	timetbl.insertAdjacentText("BeforeEnd"," ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_STime' value='00:00'>"));
 	timetbl.insertAdjacentText("BeforeEnd"," ~ �����Ͻ� ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Eday' value=''>"));
	timetbl.insertAdjacentText("BeforeEnd"," ");
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_ETime' value='00:00'>"));
	timetbl.insertAdjacentElement("BeforeEnd",document.createElement("<input type='hidden' name='lecOption' value=''>"));
}

//�൵ ���� �˾�â
function popmap(){
	popwin = window.open('/academy/lecture/lib/pop_lec_mapimg.asp','popMap','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//���� ���� �ҷ����� �˾�(���� ���¿��� �ҷ�����)
function PopOldLectureList(frm){

	popwin = window.open('/academy/lecture/lib/pop_lec_list.asp?lecturer='+ frm.lecturer_id.value ,'Listwin','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}


function InsertCate(){
	popwin = window.open('/academy/lecture/lib/pop_lec_cate.asp' ,'Listwin','width=370,height=30,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//-->
</script>

<table width="800" border="0" align="center"  class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="lecfrm" method="post" action="<%=UploadImgFingers%>/linkweb/doFingerLecture.asp">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="oldidx" value="">
<input type="hidden" name="hidweclass" value="<%=weclass%>">
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">�����ڵ�</td>
		<td width="500" bgcolor="#FFFFFF" align="left">
			<input type="button" value="�������¿��� �ҷ�����" onclick="PopOldLectureList(lecfrm);">&nbsp;&nbsp;<%=classtype%>
		</td>
	</tr>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���� ����</td>
		<td bgcolor="#FFFFFF" align="left">
			<select name="lec_gubun">
				<option value="0">�Ϲ�</option>
				<option value="1">��ü</option>
			</select>
		</td>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">Ŭ���� ����</td>
		<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD1","") %></td>
	</tr>
		<tr align="center" bgcolor="#DDDDFF">
		<td width="80">��� ����</td>
		<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD3","")%></td>
	</tr>
	</tr>
		<tr align="center" bgcolor="#DDDDFF">
		<td width="80">��� ����</td>
		<td bgcolor="#FFFFFF" align="left">
			<select name="classlevel">
				<option value="">::����::</option>
				<option value="1">�ʱ�</option>
				<option value="2">�߱�</option>
				<option value="3">���</option>
			</select>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">ī�װ�����(New)</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="hidden" name="code_large" value="">
			<input type="hidden" name="code_mid" value="">
			<input type="text" name="large_name" value="" readonly size="20"  class="text_ro">
			<input type="text" name="mid_name" value="" readonly size="20"  class="text_ro">
			<input type="button" value="ī�װ� ����" onclick="InsertCate()">
		</td>
	</tr>
    <tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
    	<td width="80">������</td>
    	<td bgcolor="#FFFFFF" align="left">
    		<select name="lecjgubun" >
    		    <option value="">����
    		    <option value="0">�⺻(���� ��õ¡�� ����)
    		    <option value="1">����������
    		</select>
    	</td>
    </tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">���¿�</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="7" size="7" name="lec_date" value="">
			(�Է¿���:2016-08)
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���¸�</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="lec_title" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">����</td>
		<td bgcolor="#FFFFFF" align="left">
		<% SelectLecturerId() %> (���縮��Ʈ�� ���»�뿩�θ� �����ϼž� ���ɴϴ�.)
			<input type="hidden" name="lecturer_id" value="">
			<input type="hidden" name="lecturer_name" value="">
			<input type="hidden" name="lecturer_regdate" value="">
		</td>
	</tr>
	
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">�൵</td>
		<td bgcolor="#FFFFFF" align="left">
		    <input class="text_ro" type="text" maxlength="3" size="3" name="map_idx" value="" readOnly >
			<input class="input_a" type="text" maxlength="128" size="64" name="lec_mapimg" value="">
			<input type="button" value="�൵ã��" onclick="javascript:popmap();">
		</td>
	</tr>
    <tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">�������</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="lec_space" value="">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="5"></td>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���� ����</td>
		<td bgcolor="#FFFFFF" align="left">
			<table width="600" border="0" class="a" cellpadding="0" cellspacing="0">
				<tr>
					<td width="60">&nbsp;</td>
					<td align="Center">������</td>
					<td align="Center" width="10">X</td>
					<td align="Center">�⺻����</td>
					<td align="Center" width="10">=</td>
					<td align="Left">���԰�</td>
				</tr>
				<tr>
				    <td>&nbsp;</td>
					<td align="Center"><input class="input_a" type="text" maxlength="10" size="8" name="lec_cost" value="0"></td>
					<td align="Center">&nbsp;</td>
					<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="margin"  value="50">%</td>
					<td align="Center">&nbsp;</td>
					<td align="left">
						<input class="input_a" type="text" maxlength="10" size="8" name="buying_cost"  value="">
						<input type="button" value="���԰� �ڵ� ���" class="button" onclick="javascript:CalcuAuto(lecfrm);">
					</td>
				</tr>
				<tr>
				    <td height="10" colspan="6"></td>
				</tr>
				<tr>
				    <td align="Center"></td>
					<td align="Center">����</td>
					<td align="Center" width="10">X</td>
					<td align="Center">�⺻����</td>
					<td align="Center" width="10">=</td>
					<td align="Left">���԰�</td>
				</tr>
				<!-- 2010 ������� ���� ���� matinclude_yn="Y"�� ���� ���� ����0�� ���� matinclude_yn="X"�� ���� -->
				<tr>
				    <td align="Center">
					    <select name="matinclude_yn" onChange="">
					    <option value="X"  >���� ����
					    <option value="C" >���� �Բ�����
					    
						<!-- <option value="N"  style='color:#999999'>���� �������(�������) -->
					    </select>
					</td>
					<td align="Center">    
					    <input class="input_a" type="text" maxlength="10" size="8" name="mat_cost" value="0">
					</td>
					<td align="Center">&nbsp;</td>
					<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="mat_margin"  value="0">%</td>
					<td align="Center">&nbsp;</td>
					<td align="left">
					    <input class="input_a" type="text" maxlength="10" size="8" name="mat_buying_cost"  value="0">
					    <input type="button" value="���԰� �ڵ� ���" class="button" onclick="javascript:CalcuAutoMaT(lecfrm);">
					</td>
				</tr>
				
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���ϸ���</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="15" size="10" name="mileage" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">��ǰ����</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="200" size="90" name="lec_attribute" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">��ǰũ��</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="200" size="90" name="lec_size" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">��ἳ��</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="mat_contents" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">Ű������</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="64" name="keyword" value="">
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td colspan="2" height="5"></td>
	</tr>

	<!--�߰� �Է� -->

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80"><%=chkiif(weclass="Y","�ִ��ο�","�����ο�")%></td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="8" size="4" name="limit_count" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">�ּ��ο�</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="8" size="4" name="min_count" value="">
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">������������</td>
		<td bgcolor="#FFFFFF" align="left">
	        <input id="reg_startday" name="reg_startday" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="reg_startday_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">������������</td>
		<td bgcolor="#FFFFFF" align="left">
	        <input id="reg_endday" name="reg_endday" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="reg_endday_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
				var CAL_Start = new Calendar({
					inputField : "reg_startday", trigger    : "reg_startday_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "reg_endday", trigger    : "reg_endday_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���� Ƚ�� <br>/ �ð�</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="8" size="4" name="lec_count" value=""> ȸ
			&nbsp;&nbsp;
			��<input class="input_a" type="text" maxlength="8" size="4" name="lec_time" value="">�ð�
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">���±Ⱓ</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="128" size="32" name="lec_period" value="">(ex : ���� �ݿ��� ���~���)
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
		<td width="80">���½ð�</td>
		<td bgcolor="#FFFFFF" align="left">
			<table width="500" class="a" border="0" cellpadding="0" cellspacing="0" >
			<tr>
				<td>
					<div class="a" id="timetbl">
						<div>
						�����Ͻ� <input class="input_a" type="text" maxlength="10" size="10" name="lec_Sday" value="<%=Date()%>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_STime" value="00:00"> ~
						�����Ͻ� <input class="input_a" type="text" maxlength="10" size="10" name="lec_Eday" value="<%=Date()%>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_ETime" value="00:00">
						<input type="hidden" name="lecOption" value="">
						</div>
					</div>
				</td>
				<td><div class="a"><input type="button" value="�ð��߰�" onclick="javascript:addtime();"></div></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���¼Ұ�</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_outline" cols="76" rows="7"></textarea>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���³���</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_contents" cols="76" rows="7"></textarea>
		</td>
	</tr>

	<!--2016-05-20 ���¿� �߰� -->
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">Ŀ��ŧ��</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_curriculum" cols="76" rows="10">[day]</textarea>
			<br><font color="red">�� ���� ������ [day] �� ������.</font>
		</td>
	</tr>

	<!--2016-05-20 ���¿� �߰�(����� ���ǻ���) -->
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">����� ���ǻ���</td>
		<td bgcolor="#FFFFFF" align="left">
			<textarea name="lec_mocaution" cols="76" rows="10"></textarea>
		</td>
	</tr>
	
	<!-- 2016-05-19 ���¿� �߰�(������url)-->
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">������URL</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" size="90" name="lec_movie" value="">
			<br>
			<font color="red">
				<!--�� ��޿� : copy embed code ���� (�� :</font><font color="blue"> //player.vimeo.com/video/102309330</font><font color="red"> ) http: ����<br>-->
				�� ��Ʃ�� : �ҽ��ڵ� ���� (�� : </font><font color="blue">http://www.youtube.com/embed/qj4rn1I_dC8 </font><font color="red">)
				�� ��Ʃ�� ������ URL���� �ƴ�!
			</font>
		</td>
	</tr>

	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">�����غ�</td>
		<td bgcolor="#FFFFFF" align="left">
			<input class="input_a" type="text" maxlength="200" size="90" name="lec_prepare" value="">
		</td>
	</tr>
	<% if (FALSE) then %>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">��Ÿ����</td>
		<td bgcolor="#FFFFFF" align="left">
			<% 
				'�������� �ʺ�� ���̸� ����
				dim editor_width, editor_height, brd_content
				editor_width = "100%"
				editor_height = "250"
			%>
			<!-- INCLUDE Virtual="/lib/util/editor.asp" -->
			<input type="hidden" name="lec_etccontents" value="">
			<font color="#8c7301">
			<br>��1. ���ܳ����� - ���� (Enter Key)
			<br>��2. �೪���� - ����Ʈ + ���� (Shift + Enter Key)
			</font>
		</td>
	</tr>
    <% end if %>
    <tr align="center" bgcolor="#DDDDFF">
		<td width="80">��Ÿ����</td>
		<td bgcolor="#FFFFFF" align="left">
		    <textarea name="lec_etccontents" cols="76" rows="10"></textarea>
		</td>
	</tr>
	    
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">��������</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="radio" name="reg_yn" value="Y">Y
			<input type="radio" name="reg_yn" value="N" checked>N
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">���ÿ���</td>
		<td bgcolor="#FFFFFF" align="left">
			<input type="radio" name="disp_yn" value="Y">Y
			<input type="radio" name="disp_yn" value="N" checked>N
		</td>
	</tr>
	
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="2">
			<span id="imagetag"  style="display:none">
			�̹��� ��뿩��   :  <input type="checkbox" name="image_saveas_yn">(�����̹����� ����ϰ��� �Ҷ� üũ�� �ּ���)
			</span>
		</td>
	</tr>
	
	<tr align="center" bgcolor="#DDDDFF">
		<td colspan="2">
			<input  type="button" value="����" onclick="javascript:frmsub(lecfrm);">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="���">
		</td>
	</tr>
	
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->