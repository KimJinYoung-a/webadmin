<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ�
' History : 2009.04.07 ������ ����
'			2010.05.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim lec_idx ,hidweclass ,chgstyle,classtype
	lec_idx= RequestCheckvar(request("lec_idx"),10)

dim onelec
set onelec = new CLecture
	onelec.FRectidx=lec_idx
	onelec.GetOneLecture

dim lectime,i,j,oldOpCd
set lectime = new CLectime
	lectime.getlectime lec_idx

dim oLectoption
set oLectoption = new CLectOption
	oLectoption.FRectidx = lec_idx
	
	if lec_idx<>"" then
		oLectoption.GetLectOptionInfo
	end if

if (onelec.FOneItem.isWeClass) then
    hidweclass = "Y"
else
    hidweclass = "N"
end if

If hidweclass = "Y" Then
	chgstyle = "style='display:none;'"
	classtype = "<font color='red'><strong>WeClass ��ü ���� ����</storng></font>"
End If 

public Sub SelectLecturerId(byval lecturer_id)
	dim sqlStr,i
	
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
			response.write "<option value='" & db2html(rsAcademyget("lecturer_id")) & "," & db2html(rsAcademyget("lecturer_name")) & "," & rsAcademyget("lec_margin") & "," & rsAcademyget("mat_margin") & "," & left(rsAcademyget("regdate"),10) & "' "&CHKIIF(lecturer_id=(rsAcademyget("lecturer_id")),"selected","") &">" & db2html(rsAcademyget("lecturer_id")) & "(" & db2html(rsAcademyget("lecturer_name")) & ") - "&db2html(rsAcademyget("brandName"))&"</option>"
		rsAcademyget.movenext
		next
			response.write "</select>"
	end if
    rsAcademyget.Close
end Sub
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>

function popLecDateEdit1(lec_idx){
	var popwin = window.open('popLecOptionEdit.asp?lec_idx='+lec_idx,'popLecDateEdit','width=700,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function frmsub(frm){

	//var frm=document.lecfrm;

	if (frm.CateCD1.value.length < 1){
		alert('Ŭ���� ������ ������ �ּ���');
		frm.CateCD1.focus();
		return;
	}

	if (frm.CateCD3.value.length < 1){
		alert('��� ������ ������ �ּ���');
		frm.CateCD3.focus();
		return;
	}
	<% if hidweclass="" or hidweclass="N" then%>
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
	
	if ((frm.lec_count.value.length < 1)||(!IsDigit(frm.lec_count.value))){
	    alert('���� Ƚ���� ���ڸ� �����մϴ�.');
		frm.lec_count.focus();
		return;
	}
	
	if ((frm.lec_time.value.length < 1)||(!IsDouble(frm.lec_time.value))){
	    alert('���� �ð��� ���ڸ� �����մϴ�.');
		frm.lec_time.focus();
		return;
	}
	
	
    //����
    
    
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
    
	// ��Ÿ���� �����Ϳ��� ������ �Է�
//	if (sector_1.chk==0){
//		frm.lec_etccontents.value = editor.document.body.innerHTML;
//	} else if(sector_1.chk!=3){
//		frm.lec_etccontents.value = editor.document.body.innerText;
//	}

    if (confirm('�����Ͻðڽ��ϱ�?')){
	    frm.submit();
	}
}

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
	//imileage = parseInt(isellcash*0.01) ;
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


function showtime(){
	var frm=timetbl.style;
	if(frm.display=="none"){
		frm.display='';
	} else {
		frm.display='none';
	}
}

//���½ð� �߰��� input box �߰� 
function addtime(tgt,opCd){
	var tfrm = document.all[tgt];
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<div>"));
	tfrm.insertAdjacentText("BeforeEnd","�����Ͻ� ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Sday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_STime' value='00:00'>"));
 	tfrm.insertAdjacentText("BeforeEnd"," ~ �����Ͻ� ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Eday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_ETime' value='00:00'>"));
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input type='hidden' name='lecOption' value='" + opCd + "'>"));
}

//Ŀ��ŧ�� �߰��� input box �߰� 
function addtime(tgt,opCd){
	var tfrm = document.all[tgt];
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<div>"));
	tfrm.insertAdjacentText("BeforeEnd","�����Ͻ� ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Sday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_STime' value='00:00'>"));
 	tfrm.insertAdjacentText("BeforeEnd"," ~ �����Ͻ� ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='10' size='10' name='lec_Eday' value=''>"));
	tfrm.insertAdjacentText("BeforeEnd"," ");
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input class='input_a' type='text' maxlength='5' size='5' name='lec_ETime' value='00:00'>"));
	tfrm.insertAdjacentElement("BeforeEnd",document.createElement("<input type='hidden' name='lecOption' value='" + opCd + "'>"));
}

function popmap(){
	popwin = window.open('/academy/lecture/lib/pop_lec_mapimg.asp','popMap','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//ī�װ�
function InsertCate(){
	popwin = window.open('/academy/lecture/lib/pop_lec_cate.asp' ,'Listwin','width=370,height=30,scrollbars=yes,resizable=yes');
	popwin.focus();
}

<% ''/������ �ű�ī�װ� �̸� ����� ���� �ӽ�. �������� ���� %>
function tmpInsertCate(){
	popwin = window.open('/academy/lecture/lib/pop_lec_cate_tmp.asp' ,'Listwin','width=400,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<form name="lecfrm" method="post" action="<%=UploadImgFingers%>/linkweb/doFingerLecture.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modi">
<input type="hidden" name="hidweclass" value="<%=hidweclass%>">
<input type="hidden" name="idx" value="<%=onelec.FOneItem.Fidx %>">

<table width="800" border="0" align="center"  class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">�����ڵ�</td>
	<td width="500" bgcolor="#FFFFFF" align="left"><%=onelec.FOneItem.Fidx %>&nbsp;&nbsp;<%=classtype%></td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���� ����</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="lec_gubun">
			<option value="0" <%=chkiif(onelec.FOneItem.Flec_gubun = "0","selected","")%>>�Ϲ�</option>
			<option value="1" <%=chkiif(onelec.FOneItem.Flec_gubun = "1","selected","")%>>��ü</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">Ŭ���� ����</td>
	<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD1",onelec.FOneItem.FCateCD1) %></td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">��� ����</td>
	<td bgcolor="#FFFFFF" align="left"><%=makeCateSelectBox("CateCD3",onelec.FOneItem.FCateCD3)%></td>
</tr>
</tr>
	<tr align="center" bgcolor="#DDDDFF">
	<td width="80">��� ����</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="classlevel">
			<option value="" <%=chkiif(onelec.FOneItem.Fclasslevel = "","selected","")%>>::����::</option>
			<option value="1" <%=chkiif(onelec.FOneItem.Fclasslevel = "1","selected","")%>>�ʱ�</option>
			<option value="2" <%=chkiif(onelec.FOneItem.Fclasslevel = "2","selected","")%>>�߱�</option>
			<option value="3" <%=chkiif(onelec.FOneItem.Fclasslevel = "3","selected","")%>>���</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">ī�װ�����</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="hidden" name="code_large" value="<%=onelec.FOneItem.Fcode_large%>">
		<input type="hidden" name="code_mid" value="<%=onelec.FOneItem.Fcode_mid%>">
		<input type="text" name="large_name" value="<%=onelec.FOneItem.Fcode_large_nm%>" readonly size="20"  class="text_ro">
		<input type="text" name="mid_name" value="<%=onelec.FOneItem.Fcode_mid_nm%>" readonly size="20"  class="text_ro">
		<input type="button" value="ī�װ� ����" onclick="InsertCate();" class="button">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">������</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="lecjgubun" >
		    <option value="">����
		    <option value="0" <%=chkiif(onelec.FOneItem.Flecjgubun = "0","selected","")%>>�⺻(���� ��õ¡�� ����)
		    <option value="1" <%=chkiif(onelec.FOneItem.Flecjgubun = "1","selected","")%>>����������
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">���¿�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="7" size="7" name="lec_date" value="<%= onelec.FOneItem.Flec_date %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���¸�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="lec_title" value="<%= onelec.FOneItem.Flec_title %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">����</td>
	<td bgcolor="#FFFFFF" align="left">
	<% SelectLecturerId(onelec.FOneItem.Flecturer_id) %> (���縮��Ʈ�� ���»�뿩�θ� �����ϼž� ���ɴϴ�.)
		<input type="hidden" name="lecturer_id" value="<%= onelec.FOneItem.Flecturer_id %>">
		<input type="hidden" name="lecturer_regdate" value="<%= left(onelec.FOneItem.Flecturer_regdate,10) %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">�����</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" name="lecturer_name" value="<%= onelec.FOneItem.Flecturer_name %>" size="10" maxlength="16">
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">�൵</td>
	<td bgcolor="#FFFFFF" align="left">
	    <input class="text_ro" type="text" maxlength="3" size="3" name="map_idx" value="<%= onelec.FOneItem.Fmap_idx %>" readOnly >
		<input class="input_a" type="text" maxlength="128" size="64" name="lec_mapimg" value="<%= onelec.FOneItem.Flec_mapimg %>">
		<input type="button" value="�൵ã��" onclick="javascript:popmap();" class="button">
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF" <%=chgstyle%>>
	<td width="80">�������</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="lec_space" value="<%= onelec.FOneItem.Flec_space %>">
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
				<td align="Center"><input class="input_a" type="text" maxlength="10" size="8" name="lec_cost" value="<%= onelec.FOneItem.Flec_cost %>"></td>
				<td align="Center">&nbsp;</td>
				<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="margin"  value="<%= onelec.FOneItem.Fmargin %>">%</td>
				<td align="Center">&nbsp;</td>
				<td align="left">
					<input class="input_a" type="text" maxlength="10" size="8" name="buying_cost"  value="<%= onelec.FOneItem.Fbuying_cost %>">
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
			<!-- ����
			<input type="checkbox" name="matinclude_yn" onclick="javascript:Fnmat();" <% if onelec.FOneItem.Fmatinclude_yn="Y" then response.write "checked" %>>(��������)
			<input class="input_a" type="text" maxlength="16" size="8" name="mat_cost" value="<%= onelec.FOneItem.Fmat_cost %>">
			-->
			<!-- 2010 ������� ���� ���� matinclude_yn="Y"�� ���� ���� ����0�� ���� matinclude_yn="X"�� ���� -->
			
			<tr>
			    <td align="Center">
			    <!-- option value="N" onelec.FOneItem.Fmatinclude_yn="N" ����� -->
				    <select name="matinclude_yn" onChange="">
				    <option value="X"  <%= CHKIIF(onelec.FOneItem.Fmatinclude_yn="X" or onelec.FOneItem.Fmat_cost=0,"selected","") %>>���� ����
				    <option value="C" <%= CHKIIF(onelec.FOneItem.Fmatinclude_yn="C","selected","") %> >���� �Բ�����
				    
					<!-- <option value="N" <%= CHKIIF(onelec.FOneItem.Fmatinclude_yn="N" or (onelec.FOneItem.Fmatinclude_yn="N" and onelec.FOneItem.Fmat_cost>0),"selected","") %> style='color:#999999'>���� �������(�������) -->
				    </select>
				</td>
				<td align="Center">    
				    <input class="input_a" type="text" maxlength="10" size="8" name="mat_cost" value="<%= onelec.FOneItem.Fmat_cost %>">
				</td>
				<td align="Center">&nbsp;</td>
				<td align="Center"><input class="input_a" type="text" maxlength="8" size="4" name="mat_margin"  value="<%= onelec.FOneItem.Fmat_margin %>">%</td>
				<td align="Center">&nbsp;</td>
				<td align="left">
				    <input class="input_a" type="text" maxlength="10" size="8" name="mat_buying_cost"  value="<%= onelec.FOneItem.Fmat_buying_cost %>">
				    <input type="button" value="���԰� �ڵ� ���" class="button" onclick="javascript:CalcuAutoMaT(lecfrm);">
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���ϸ���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="15" size="10" name="mileage" value="<%= onelec.FOneItem.Fmileage %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">��ǰ����</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="200" size="90" name="lec_attribute" value="<%= onelec.FOneItem.Flec_attribute %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">��ǰũ��</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="200" size="90" name="lec_size" value="<%= onelec.FOneItem.Flec_size %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">��ἳ��</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="mat_contents" value="<%= onelec.FOneItem.Fmat_contents %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">Ű������</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="128" size="64" name="keyword" value="<%= onelec.FOneItem.Fkeyword %>">
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td colspan="2" height="5"></td>
</tr>

<!--�߰� �Է� -->

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">�� �����ο�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="limit_count" value="<%= onelec.FOneItem.Flimit_count %>" readonly style="background-color='#EEEEEE'">
		�� �����ο��� ������ ����(�ɼ�)�������� ���ּ���.
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">�� �����ο�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="limit_sold" value="<%= onelec.FOneItem.Flimit_sold %>" readonly style="background-color='#EEEEEE'">
		�� �����ο��� ������ ����(�ɼ�)�������� ���ּ���.
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���´�<br>�ּ��ο�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="min_count" value="<%= onelec.FOneItem.Fmin_count %>">
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">������������</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="reg_startday" name="reg_startday" value="<%=onelec.FOneItem.Freg_startday%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="reg_startday_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">������������</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="reg_endday" name="reg_endday" value="<%=onelec.FOneItem.Freg_endday%>" class="text" size="10" maxlength="10" />
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
	<td width="80">���� Ƚ�� /<br> �ð�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="8" size="4" name="lec_count" value="<%= onelec.FOneItem.Flec_count %>"> ȸ
		&nbsp;&nbsp;
		��<input class="input_a" type="text" maxlength="8" size="4" name="lec_time" value="<%= onelec.FOneItem.Flec_time %>">�ð�
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">��������<br><input type="button" value="��������" onclick="popLecDateEdit1('<%= lec_idx %>');" class="button"></td>
	<td bgcolor="#FFFFFF" align="left">
	<!-- �����ÿ��� �ɼ� ���.. -->
	<table width="100%" border="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        	<td width="40">�ڵ�</td>
        	<td width="100">�����Ⱓ</td>
        	<td >�����Ͻ�</td>
        	<td width="30">���<br>����</td>
        	<td width="30">����</td>
        	<td width="30">��û<br>�ο�</td>
        	<td width="30">����<br>�ο�</td>
        	<td width="40">����<br>����</td>
        </tr>
	<% for i=0 to oLectoption.FResultCount -1 %>
	
	    <tr align="center" bgcolor="<%= ChkIIF(oLectoption.FItemList(i).Fisusing="Y","#FFFFFF","#DDDDDD") %>">
        	<td><%= oLectoption.FItemList(i).FlecOption %></td>
        	<td>
        		����: <%= FormatDateTime(oLectoption.FItemList(i).FRegStartDate,2) %><br>
        		����: <%= FormatDateTime(oLectoption.FItemList(i).FRegEndDate,2) %>
        	</td>
        	<td align="left">
        		<%=oLectoption.FItemList(i).FlecOptionName%><br>
        		<%= FormatDateTime(oLectoption.FItemList(i).FlecStartDate,2) %>&nbsp;
        		<%= FormatDateTime(oLectoption.FItemList(i).FlecStartDate,4) %>~
        		<%= FormatDateTime(oLectoption.FItemList(i).FlecEndDate,4) %>
        	</td>
        	<td>
        	    <%= oLectoption.FItemList(i).Fisusing %>
        	</td>
        	<td><%= oLectoption.FItemList(i).Flimit_count %></td>
        	<td><%= oLectoption.FItemList(i).Flimit_sold %></td>
        	<td><%= oLectoption.FItemList(i).Flimit_count-oLectoption.FItemList(i).Flimit_sold %></td>
        	<td><% if oLectoption.FItemList(i).IsOptionSoldOut then %><font color="red">����</font><% end if %></td>
        </tr>
	<% next %>
	</table>
	<!-- lec_period ������...
		<input class="input_a" type="text" maxlength="128" size="40" name="lec_period" value="<%= onelec.FOneItem.Flec_period %>" readonly style="background-color='#EEEEEE'">(ex : ���� �ݿ��� ���~���)
	    -->
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���½ð�</td>
	<td bgcolor="#FFFFFF" align="left">
		<table width="100%" class="a" border="0" cellpadding="2" cellspacing="2" >
		<%
			for i = 0 to lectime.FResultcount -1
				if oldOpCd<>lectime.FlecOption(i) then
					j=j+1
		%>
		<tr align="center">
			<td width="50" bgcolor="#E8E8E8"><%=lectime.FlecOption(i)%></td>
			<td bgcolor="#F2F2F2"><div class="a" id="timetbl<%=j%>">
		<%
				end if
				oldOpCd = lectime.FlecOption(i)
		%>
					<div>
					�����Ͻ� <input class="input_a" type="text" maxlength="10" size="10" name="lec_Sday" value="<% = formatdatetime(lectime.FStartDate(i),2) %>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_STime" value="<% = formatdatetime(lectime.FStartDate(i),4) %>"> ~
					�����Ͻ� <input class="input_a" type="text" maxlength="10" size="10" name="lec_Eday" value="<% = formatdatetime(lectime.FEndDate(i),2) %>"> <input class="input_a" type="text" maxlength="5" size="5" name="lec_ETime" value="<% = formatdatetime(lectime.FEndDate(i),4) %>">
					<input type="hidden" name="lecOption" value="<%=lectime.FlecOption(i)%>">
					</div>
			<% if (i>=(lectime.FResultcount-1) ) or (i<lectime.FResultcount and oldOpCd<>lectime.FlecOption(i+1)) then %>
				</div>
			</td>
			<td width="110" bgcolor="#F2F2F2"><div class="a"><input type="button" value="�ð��߰� #<%=j%>" onclick="javascript:addtime('timetbl<%=j%>','<%=lectime.FlecOption(i)%>');" class="button"></div></td>
		</tr>
			<%
				end if
			%>
		<% next %>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���¼Ұ�</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_outline" cols="76" rows="7"><%= onelec.FOneItem.Flec_outline %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���³���</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_contents" cols="76" rows="7"><%= onelec.FOneItem.Flec_contents %></textarea>
	</td>
</tr>

<!--2016-05-20 ���¿� �߰� -->
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">Ŀ��ŧ��</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_curriculum" cols="76" rows="10"><%=chkiif(onelec.FOneItem.Flec_curriculum = "","[day]1",onelec.FOneItem.Flec_curriculum)%></textarea>
		<br><font color="red">�� ���� ������ [day] �� ������.</font>
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">�����غ�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" maxlength="200" size="90" name="lec_prepare" value="<%= onelec.FOneItem.Flec_prepare %>">
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

			brd_content = onelec.FOneItem.Flec_etccontents
			if inStr(brd_content,"<br>")=0 and inStr(brd_content,"<P>")=0 then
				brd_content = replace(brd_content,vbCrLf,"<br>")
			end if
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
	    <textarea name="lec_etccontents" cols="76" rows="10"><%= onelec.FOneItem.Flec_etccontents %></textarea>
	</td>
</tr>    
<!--2016-05-20 ���¿� �߰�(����� ���ǻ���) -->
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">����� ���ǻ���</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="lec_mocaution" cols="76" rows="10"><%= onelec.FOneItem.Flec_mocaution %></textarea>
	</td>
</tr>

<!-- 2016-05-19 ���¿� �߰�(������url)-->
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">������URL</td>
	<td bgcolor="#FFFFFF" align="left">
		<input class="input_a" type="text" size="90" name="lec_movie" value="<%= onelec.FOneItem.Flec_movie %>">
		<br>
		<font color="red">
			<!--�� ��޿� : copy embed code ���� (�� :</font><font color="blue"> //player.vimeo.com/video/102309330</font><font color="red"> ) http: ����<br>-->
			�� ��Ʃ�� : �ҽ��ڵ� ���� (�� : </font><font color="blue">http://www.youtube.com/embed/qj4rn1I_dC8 </font><font color="red">)
			�� ��Ʃ�� ������ URL���� �ƴ�!
		</font>
	</td>
</tr>

<tr align="center" bgcolor="#DDDDFF">
	<td width="80">��������</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="radio" name="reg_yn" value="Y" <% if onelec.FOneItem.Freg_yn="Y" then response.write "checked" %>>Y
		<input type="radio" name="reg_yn" value="N" <% if onelec.FOneItem.Freg_yn="N" then response.write "checked" %>>N
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td width="80">���ÿ���</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="radio" name="disp_yn" value="Y" <% if onelec.FOneItem.Fdisp_yn="Y" then response.write "checked" %>>Y
		<input type="radio" name="disp_yn" value="N" <% if onelec.FOneItem.Fdisp_yn="N" then response.write "checked" %>>N
	</td>
</tr>
<tr align="center" bgcolor="#DDDDFF">
	<td colspan="2">
		<input  type="button" value="����" onclick="javascript:frmsub(lecfrm);" class="button">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="���" class="button">
	</td>
</tr>
</table>

</form>
<%
set onelec = nothing
set lectime = nothing
set oLectoption = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->