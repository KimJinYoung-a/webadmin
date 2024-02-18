<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ΰŽ� ��ī���� �Ű��� ���,���� �˾�
' History : 2016-03-03 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/academy/magazineCls.asp" -->
<%
	Dim idx, oMagazine, sqlStr
	Dim classcode, classcodearr, classcodecnt, catecode
	Dim listimg, viewimg1, viewimg2, viewimg3, viewno, state, startdate, viewtitle, viewtext1, viewtext2, viewtext3, videourl, isusing

	idx = RequestCheckvar(request("idx"),10)
	set oMagazine = new CMagazineContents
		 oMagazine.FRectIdx = idx

		if idx <> "" Then
			oMagazine.GetOneRowMagaZineContent()
			if oMagazine.FResultCount > 0 then
				state		= oMagazine.FOneItem.Fstate
				viewno		= oMagazine.FOneItem.Fviewno
				listimg	= oMagazine.FOneItem.Flistimg
				viewimg1	= oMagazine.FOneItem.Fviewimg1
				viewimg2	= oMagazine.FOneItem.Fviewimg2
				viewimg3	= oMagazine.FOneItem.Fviewimg3
				viewtext1	= oMagazine.FOneItem.Fviewtext1
				viewtext2	= oMagazine.FOneItem.Fviewtext2
				viewtext3	= oMagazine.FOneItem.Fviewtext3
				videourl	= oMagazine.FOneItem.Fvideourl
				catecode	= oMagazine.FOneItem.Fcatecode
				viewtitle	= oMagazine.FOneItem.Fviewtitle
				startdate	= oMagazine.FOneItem.Fstartdate
				classcode	= oMagazine.FOneItem.Fclasscode
				isusing	= oMagazine.FOneItem.Fisusing
			end if
		end if
	set oMagazine = Nothing

	if idx = 0 then
		sqlStr = " delete from db_academy.dbo.tbl_academy_magazine_keyword where vidx = 0 "
		dbACADEMYget.Execute sqlStr
	end if

	''Ŭ�����ڵ�(ī�װ��ڵ�) 0~5�� ����
	classcodearr = split(classcode,",")
	classcodecnt = ubound(classcodearr)+1
	if isusing = "" then isusing = "Y"
%>
<script type="text/javascript">
	//''jsPopCal : �޷� �˾�
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//''��ϵ� �̹��� ����
	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	//''�̹��� ���
	function jsSetImg(sImg, sName, sSpan){
		document.domain ="10x10.co.kr";

		var winImg;
		winImg = window.open('/academy/magazine/lib/pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	//''�±� ����
	function jsTagview(idx){
		var poptag;
		poptag = window.open('/academy/magazine/lib/pop_tagReg.asp?idx='+idx,'poptag','width=300,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}

	//''ī�װ� ����
//	function jsCatecodeview(idx){
//		var poptag;
//		poptag = window.open('/academy/magazine/lib/pop_catecodeReg.asp','popcatecode','width=300,height=400,scrollbars=yes,resizable=yes');
//		poptag.focus();
//	}

	//����
	function subcheck(){
		var frm=document.inputfrm;

		if(!frm.catecode.value){
			alert("������ �������ּ���");
			frm.catecode.focus();
			return;
		}

		if (!frm.viewno.value){
			alert('No.�� ������ּ���');
			frm.viewno.focus();
			return;
		}
        if(isNaN(frm.viewno.value) == true) {
            alert("���ڸ� �Է� �����մϴ�.");
            frm.viewno.focus();
            return false;
        }
	    
		if (!frm.viewtitle.value){
			alert('������ ������ּ���');
			frm.viewtitle.focus();
			return;
		}
		if(!frm.state.value){
			alert("���¸� �������ּ���");
			frm.state.focus();
			return;
		}

		if (!frm.startdate.value){
			alert('�������� ������ּ���');
			frm.startdate.focus();
			return;
		}
//		if (!frm.viewtext1.value){
//			alert('�󼼳����� ������ּ���');
//			frm.viewtext1.focus();
//			return;
//		}

	    for(var i=0;i<5;i++){
	        if(isNaN(frm['classcode'+(i+1)].value) == true) {
	            alert("���ڸ� �Է� �����մϴ�.");
	            frm['classcode'+(i+1)].focus();
	            return false;
	        }
	    }
		frm.classcode.value = frm.classcode1.value+","+frm.classcode2.value+","+frm.classcode3.value+","+frm.classcode4.value+","+frm.classcode5.value;
		frm.submit();
	}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="magazineProc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="listimg" value="<%= listimg %>">
<input type="hidden" name="viewimg1" value="<%= viewimg1 %>">
<input type="hidden" name="viewimg2" value="<%= viewimg2 %>">
<input type="hidden" name="viewimg3" value="<%= viewimg3 %>">
<input type="hidden" name="classcode" value="">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Academy&gt;&gt;�Ű��� ���/����</b></font>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
	<td bgcolor="#FFFFFF">
		<b><%=idx%></b>
	</td>
</tr>
<% End If %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<% DrawMagazineGubun "catecode" , catecode ,"" %>
		<!--<input type="button" name="btnviewImg" value="ī�װ� ����" onClick="jsCatecodeview();" class="button"/>-->
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">No.</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewno" value="<%=viewno%>" size="10"/>�� ���ڰ� Ŭ���� �켱 ǥ�� �˴ϴ�. ��
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewtitle" value="<%=viewtitle%>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %> �� ������ �ؼ� �����Ͽ��� ������ =< ���� �̾�߸� ������ �˴ϴ�.
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
   		<% IF state = "9" THEN %>
   			<%= startdate %><input type="hidden" name="startdate" size=20 maxlength=10 value="<%= startdate %>"/>
   		<% ELSE %>
   			<input type="text" name="startdate" size=20 maxlength=10 value="<%= startdate %>" onClick="jsPopCal('startdate');"  style="cursor:pointer;"/>
   		<% END IF %>
		<font color="red">��Ŭ���� �޷¿��� ����</font>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����Ʈ�̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg" value="�̹������" onClick="jsSetImg('<%=listimg%>','listimg','listimgdiv')" class="button"/>
		<div id="listimgdiv" style="padding: 5 5 5 5">
			<% IF listimg <> "" THEN %>
				<img src="<%=listimg%>" border="0" height=100 onclick="jsImgView('<%=listimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('listimg','listimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a>
			<% END IF %>
		</div>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�󼼳���1</td>
	<td bgcolor="#FFFFFF">
		<textarea name="viewtext1" rows="8" cols="50"><%=viewtext1%></textarea>
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹���1</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg1" value="�̹������" onClick="jsSetImg('<%=viewimg1%>','viewimg1','viewimgdiv1')" class="button"/>
		<div id="viewimgdiv1" style="padding: 5 5 5 5">
			<% IF viewimg1 <> "" THEN %>
				<img src="<%=viewimg1%>" border="0" height=100 onclick="jsImgView('<%=viewimg1%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('viewimg1','viewimgdiv1');"><img src="/images/icon_delete2.gif" border="0"/></a>
			<% END IF %>
		</div>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�󼼳���2</td>
	<td bgcolor="#FFFFFF">
		<textarea name="viewtext2" rows="8" cols="50"><%=viewtext2%></textarea>
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹���2</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg2" value="�̹������" onClick="jsSetImg('<%=viewimg2%>','viewimg2','viewimgdiv2')" class="button"/>
		<div id="viewimgdiv2" style="padding: 5 5 5 5">
			<% IF viewimg2 <> "" THEN %>
				<img src="<%=viewimg2%>" border="0" height=100 onclick="jsImgView('<%=viewimg2%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('viewimg2','viewimgdiv2');"><img src="/images/icon_delete2.gif" border="0"/></a>
			<% END IF %>
		</div>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�󼼳���3</td>
	<td bgcolor="#FFFFFF">
		<textarea name="viewtext3" rows="8" cols="50"><%=viewtext3%></textarea>
	</td>
</tr>

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹���3</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg3" value="�̹������" onClick="jsSetImg('<%=viewimg3%>','viewimg3','viewimgdiv3')" class="button"/>
		<div id="viewimgdiv3" style="padding: 5 5 5 5">
			<% IF viewimg3 <> "" THEN %>
				<img src="<%=viewimg3%>" border="0" height=100 onclick="jsImgView('<%=viewimg3%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('viewimg3','viewimgdiv3');"><img src="/images/icon_delete2.gif" border="0"/></a>
			<% END IF %>
		</div>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ũ</td>
	<td bgcolor="#FFFFFF">
		<textarea name="videourl" rows="5" cols="50"><%=videourl%></textarea><br/><br/>
		<font color="red">
			�� ������ : �ҽ��ڵ� ���� (�� : </font><font color="blue">http://www.youtube.com/embed/qj4rn1I_dC8 </font><font color="red">)<br>
			�� ��޿� : copy embed code ���� (�� :</font><font color="blue"> //player.vimeo.com/video/102309330</font><font color="red"> ) http: ����
		</font>
	</td>
</tr>
<tr align="left" bgcolor="<%= adminColor("topbar") %>" >
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� �����ڵ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="classcode1" value="<% if classcodecnt > 0 then response.write classcodearr(0) else response.write "" end if %>" size="10" maxlength = "10" />
		<input type="text" class="text" name="classcode2" value="<% if classcodecnt > 1 then response.write classcodearr(1) else response.write "" end if %>" size="10" maxlength = "10" />
		<input type="text" class="text" name="classcode3" value="<% if classcodecnt > 2 then response.write classcodearr(2) else response.write "" end if %>" size="10" maxlength = "10" />
		<input type="text" class="text" name="classcode4" value="<% if classcodecnt > 3 then response.write classcodearr(3) else response.write "" end if %>" size="10" maxlength = "10" />
		<input type="text" class="text" name="classcode5" value="<% if classcodecnt > 4 then response.write classcodearr(4) else response.write "" end if %>" size="10" maxlength = "10" />
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�������� �±�</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg" value="�±� ����" onClick="jsTagview('<%=idx%>')" class="button"/><br/>
		���±װ����� �˾����� ���� �մϴ� ���� ��� ���ּ���.��<br/>
		������ ������ �±ס�<br/>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
	<td colspan="2">
		<input type="radio" name="isusing" value="Y" <%=chkIIF(isusing="Y","checked","")%>/>����� &nbsp;&nbsp;&nbsp; 
		<input type="radio" name="isusing" value="N" <%=chkIIF(isusing="N","checked","")%>/>������
	</td>
</tr>

<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="window.close();"/>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->