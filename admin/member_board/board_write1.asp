<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���
' History : 2011.02.23 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim g_MenuPos, writer
	IF application("Svr_Info")="Dev" THEN
		g_MenuPos   = "1288"		'### �޴���ȣ ����.
	Else
		g_MenuPos   = "1304"		'### �޴���ȣ ����.
	End If

	Dim cBoard
	Dim sBrd_Id, sBrd_Name, sBrd_Regdate
	Dim part_sn, job_sn, posit_sn
	Dim brd_content, arrFileList
	sBrd_Id 		= session("ssBctId")
	sBrd_Name		= session("ssBctCname")
	sBrd_Regdate	= Left(now(),10)
 
	set cBoard = new Board
		cBoard.fnBoardcontent
%>
<!-- daumeditor head ------------------------->
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 <meta http-equiv="X-UA-Compatible" content="IE=10" /> 
 <link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="utf-8"/>
 <script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="utf-8"></script>
 <script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="utf-8"></script>
 <script type="text/javascript">
    var config = {
        initializedId: "",
        wrapper: "tx_trex_container",
        form: 'frm',
        txIconPath: "/lib/util/daumeditor/images/icon/editor/",
        txDecoPath: "/lib/util/daumeditor/images/deco/contents/",
        
        events: {
            preventUnload: false
        },
        sidebar: {
            attachbox: {
                show: true
            },
            attacher: {
                 image: {
                    popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
                } 
            }
            
        }
    }; 
 </script>
<!-- //daumeditor head ------------------------->
<script language="javascript">
function fTeam(str){
	if(str == "all"){
		document.getElementById('brd_team').style.display = 'none';
		for(var j=0; j<frm.part_sn.length; j++) {
			frm.part_sn[j].checked = false;
		}
	}else{
		document.getElementById('brd_team').style.display = 'block';
		//var sTeam = window.open("","","");
	}
}
function form_check(){
	var frm = document.frm;
//���� ���� //
	if(frm.brd_type.value == "")
	{
		alert("���� ������ �����ϼ���");
		return false;
	}

//�������� ����//
	var chk = 0;
	for(var i=0; i<frm.open_team.length; i++) {
		if(frm.open_team[i].checked) chk++;
	}
	if(chk == "0"){
		alert("���������� �����ϼ���");
		return false;
	}
//������ ���ý� �� ���� ����//
	if(frm.open_team[1].checked){
		var chk2 = 0;
	   if(document.all.did != undefined) {
	       chk2 =1;  
	       if(document.all.did.length != undefined) {
	            for(i=0;i<document.all.did.length;i++){
	                if(i==0){
	                     frm.arrdid.value = document.all.did[i].value;
	                }else{
	                     frm.arrdid.value = frm.arrdid.value+","+document.all.did[i].value;
	                }
	            }
	        }else{
	            frm.arrdid.value = document.all.did.value;
	        }
	    }  
 
		if(chk2 == "0"){
			alert("���� �����ϼ���");
			return false;
		}
	}
 
//������ ���ý� �������� ���� ����//
	if(frm.open_team[1].checked){
		var chk3 = 0;
		for(var k=0; k<frm.job_sn.length; k++) {
			if(frm.job_sn[k].selected){
				chk3++;
			}
		}
		if(chk3 == "0"){
			alert("��å�� �����ϼ���");
			return false;
		}
	}
//������ ���ý� ���� ���� ����//
	if(frm.open_team[1].checked){
		var chk4 = 0;
		for(var w=0; w<frm.posit_sn.length; w++) {
			if(frm.posit_sn[w].selected){
				chk4++;
			}
		}
		if(chk4 == "0"){
			alert("������ �����ϼ���");
			return false;
		}
	}
//���� �Է� ����//
	if(frm.brd_subject.value == ""){
		alert("������ �Է��ϼ���");
		frm.brd_subject.focus();
		return false;
	}

//�������� ����//
	var chk3 = 0;
	for(var k=0; k<frm.brd_fixed.length; k++) {
		if(frm.brd_fixed[k].checked) chk3++;
	}
	if(chk3 == "0"){
		alert("�������θ� �����ϼ���");
		return false;
	}
	
	  var content = Editor.getContent();
      document.getElementById("brd_content").value = content; 
//	// �̳���ͷ� ������ ���� textarea�� �Ҵ� ����
//	var strHTMLCode = fnGetEditorHTMLCode(true, 0);
//	if(strHTMLCode == ''){
//		alert("������ �Է��ϼ���");
//		return false;
//	}else{
//		frm["brd_content"].value = strHTMLCode;
//	}
//	// �̳���ͷ� ������ ���� textarea�� �Ҵ� ��
	frm.action = "board_proc.asp";
	frm.submit();
}
function fileupload()
{
	window.open('board_popupload.asp','worker','width=420,height=200,scrollbars=yes');
}
function clearRow(tdObj) {
	if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;

		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}


 //Ʈ����� Ŭ���� ���� ����
	function jsOpenClose(cValue,iValue){ 
		if(eval("document.all.divB"+cValue+iValue).style.display=="none"){
			eval("document.all.Fimg"+iValue).src = "/images/dtree/openfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tminus.png";
			eval("document.all.divB"+cValue+iValue).style.display=""; 
		}else{
			eval("document.all.Fimg"+iValue).src = "/images/dtree/closedfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tplus.png";
			eval("document.all.divB"+cValue+iValue).style.display="none"; 
		}
	}	
	
//������
    function jsSelTeam(){
        var sdid="";
        if(document.all.did!=undefined){
            if(document.all.did.length!=undefined){
                for(i=0;i<document.all.did.length;i++){
                    if(i==0){
                        sdid =  document.all.did[i].value;
                    }else{
                        sdid = sdid +"," + document.all.did[i].value;
                     }   
                }
            }else{
                sdid =  document.all.did.value;
            }
        }
        var winTeam = window.open("/admin/member_board/popselectteam.asp?did="+sdid,"popT","width=330,height=700,resizable=yes");
        winTeam.focus();
    }
    
//������
    function jsTeamDel(i){
        eval("document.all.dvSTeam"+i).outerHTML = "";
    }    	
</script>

<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td><img src="/images/icon_arrow_link.gif"></td>
	<td style="padding-top:3">&nbsp;<b>�Խñ� �ۼ�</b></td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<form name="frm"  method="post">
<input type = "hidden" name = "mode" value = "add">
<input type = "hidden" name = "brd_sn" value = "<%=cBoard.Fbrd_sn + 1%>">
<input type = "hidden" name = "arrdid" value = ""> 
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No.<%= cBoard.Fbrd_sn + 1 %></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sBrd_Name%>(<%=sBrd_Id%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �����: <%=sBrd_Regdate%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��������</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td>
						<table class="a">
						<tr>
							<td><%=fnBrdType("w", "Y", "", "")%>&nbsp;</td>
							<td><font color="RED"> �� �λ���ð��� : �ű��Ի�����ڰ���,�λ��̵�����,��Ÿ�λ����
								<br>&nbsp;&nbsp;&nbsp;&nbsp;ȸ�系�԰��ð��� : �Ի�,���,�޿� �� ȸ��� �������� ���� ����
								<br>&nbsp;&nbsp;&nbsp;&nbsp;�Ϲ� �ȳ� ���� : ��� �׸� �� ����
								</font></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��������</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td>
						<label><input type="radio" name="open_team" value="Y" checked onclick="fTeam('all');">��ü����</label>&nbsp;&nbsp;&nbsp;
						<label><input type="radio" name="open_team" value="N" onclick="fTeam('team');">������</label>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30" id = "brd_team" style="display:none">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">������</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<!--% call DrawPartinfoCombo("part_sn", part_sn,"") %-->
				     <div style="padding:5 5 5 5px;"><input type="button" class="button" value="������" onClick="jsSelTeam();"></div>
				     <div id="dvTeam" style="padding:5 5 5 5px;"></div>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��å</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<% call DrawJobCombo("job_sn", job_sn) %>&nbsp;<font color="RED"> �� �ش� ��å �̻� ���� �����մϴ�. (�Ϲ��� �����Ͻø� ��� ���������ϰ� �˴ϴ�)</font>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<% call DrawPositCombo("posit_sn", posit_sn) %>&nbsp;&nbsp;&nbsp;<font color="RED">�� �ش� ���� �̻� ���� �����մϴ�.</font>
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="brd_subject" value="" size="95" maxlength="128">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 5 4 2 5">
				<textarea name="brd_content" id="brd_content" style="width: 100%; height: 490px;"></textarea>  
                <!-- daumeditor  --> 
                <script type="text/javascript">  
                    EditorCreator.convert(document.getElementById("brd_content"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                        EditorJSLoader.ready(function (Editor) {
                            new Editor(config);
                            Editor.modify({
                                content: ''
                            });
                        });
                    });
                
                </script> 
                <!-- daumeditor   -->
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">÷������</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td width="100" valign="top" style="padding:5 0 0 0">
						<input type="button" value="���Ͼ��ε�" class="button" onclick="fileupload();">
					</td>
					<td width="100%" style="padding:3 0 3 10">
						<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
						<%
						IF isArray(arrFileList) THEN
							For i =0 To UBound(arrFileList,2)
						%>
							<tr>
								<td>
									<input type='hidden' name='doc_file' value='<%=arrFileList(0,i)%>'>
									<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
									<a href='<%=arrFileList(0,i)%>' target='_blank'>
									<%=Split(Replace(arrFileList(0,i),"http://",""),"/")(4)%></a>
								</td>
							</tr>
						<%
							Next
							Response.Write "<input type='hidden' name='isfile' value='o'>"
						Else
						%>
							<tr>
								<td>
								</td>
							</tr>
						<% End If %>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<%
			If session("ssAdminLsn") <= "3" or (session("ssBctID")="gogo27") or (session("ssBctID")="jjun531") or (session("ssBctID")="choi23") Then
		%>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="brd_useynY"><input type="radio" name="brd_fixed" id="brd_useynY" value="1">Y</label>&nbsp;&nbsp;&nbsp;
				<label id="brd_useynN"><input type="radio" name="brd_fixed" id="brd_useynN" value="2">N</label><br>
				<font color = "RED"> �ذ��� ���� Y�� �����Ͻø� �Խñ��� �ֻ�ܿ� ��ġ�ϰ� �˴ϴ�.</font>
			</td>
		</tr>
		<%
			Else
		%>
		<tr height="30" style="display:none">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="brd_useynY"><input type="radio" name="brd_fixed" id="brd_useynY" value="1">Y</label>&nbsp;&nbsp;&nbsp;
				<label id="brd_useynN"><input type="radio" name="brd_fixed" id="brd_useynN" value="2" checked>N</label><br>
				<font color = "RED"> �ذ��� ���� Y�� �����Ͻø� �Խñ��� �ֻ�ܿ� ��ġ�ϰ� �˴ϴ�.</font>
			</td>
		</tr>
		<%
			End If
		%>
		</table>
	</td>
</tr>
</form>
</table>
<table width="813" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left"><a href="board_list.asp?menupos=<%=g_MenuPos%>"><img src="/images/icon_list.gif" border="0"></a></td>
	<td width="50%" align="right">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="image" src="/images/icon_confirm.gif" border="0" onclick="form_check();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
