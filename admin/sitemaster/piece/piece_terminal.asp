<% Option Explicit %>
<%
'###########################################################
' Description : piece �͹̳�(����Ʈ) ������
' Hieditor : 2017.08.28 ������ ����
'			 2017-11-29 ����ȭ �߰� / ����
'###########################################################

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/piece/piececls.asp"-->
<!-- #include virtual="/lib/util/pageformlib.asp" -->

<%
Dim gubun	'1 : ����, 2 :����, 3 : ����ƮŰ����, 4: ���, 5:ȸ������
Dim oPieceUser, loginUserId, oPieceList, i, oPieceOpening, currpage, pagesize, deal, open, research, keyword, schWord
Dim state , schdate

loginUserId = session("ssBctId")
currpage = requestcheckvar(request("page"), 20)
deal = requestcheckvar(request("deal"), 20)
open = requestcheckvar(request("open"), 20)
research = requestcheckvar(request("research"), 20)
keyword = requestcheckvar(request("keyword"), 20)
schWord = requestcheckvar(request("schWord"), 500)
state = requestcheckvar(request("state"), 1)

'// ������ �˻�
schdate = requestcheckvar(request("prevDate"), 10)

If keyword = "snum" And Not(isNumeric(schWord)) Then
	Response.write "<script>alert('��ȣ(idx) �� Ȯ�� ���ּ���');</script>"
	schWord = ""
End If

If Trim(currpage)="" Then
	currpage = "1"
End If
pagesize = 30

'// ���� ���� �����ڰ� piece�� ��ϵ� ���������� Ȯ���Ѵ�.
set oPieceUser = new Cgetpiece
	oPieceUser.FRectadminid = loginUserId
	oPieceUser.adminPieceUser()

'// ������ �����͸� �����´�.
set oPieceOpening = new Cgetpiece
	oPieceOpening.getPieceOpening()

'// ����Ʈ�� �����´�.
set oPieceList = new Cgetpiece
	oPieceList.FRectcurrpage = currpage
	oPieceList.FRectpagesize = pagesize
	If Trim(research)="on" Then
		oPieceList.FRectDeal = deal
		oPieceList.FRectOpen = open
		oPieceList.FRectkeyword = keyword
		oPieceList.FRectSchword = schWord
		oPieceList.FRectState = state
		oPieceList.FRectStartdate = schdate
	End If
	oPieceList.GetpieceList()

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script language="JavaScript" src="/js/xl.js"></script>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type='text/javascript'>
document.domain = "10x10.co.kr";

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;
	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}
	ret = confirm(msg);
	if (ret) {
		frm.submit();
	}
}

// ������ �Է�/����
function fnPieceUseract()
{
	// nickname���� �־���.
	$("#frmnickname").val(escape($("#usernickname").val()));
	// occupation���� �־���
	$("#frmoccupation").val($("#selectOccupation").val());

	if ($("#selectOccupation").val()=="A")
	{
		alert("������ �������ּ���.");
		return;
	}
	if ($("#usernickname").val()=="")
	{
		alert("�г����� �Է����ּ���.");
		return;
	}

	<% if trim(oPieceUser.FoneUser.Fnickname)<>"" then %>
		$("#frmmode").val("upd");
	<% else %>
		$("#frmmode").val("ins");
	<% end if %>

	$.ajax({
		type:"GET",
		url:"/admin/sitemaster/piece/act_pieceUser.asp",
		data:$("#frmpieceUser").serialize(),
		async:false,
		cache:true,
		success : function(Data, textStatus, jqXHR){
			if (jqXHR.readyState == 4) {
				if (jqXHR.status == 200) {
					if(Data!="") {
						res = Data.split("||");
						if (res[0]=="OK")
						{
							if (res[1]=="2")
							{
								alert("�����Ϸ�");
							}
							document.location.reload();
						}
						else
						{
							errorMsg = res[1].replace(">?n", "\n");
							alert(errorMsg);
							document.location.reload();
							return false;
						}
					} else {
						alert("�߸��� ���� �Դϴ�.");
						document.location.reload();
						return false;
					}
				}
			}
		},
		error:function(jqXHR, textStatus, errorThrown){
			alert("�߸��� ���� �Դϴ�.");
			<% if false then %>
				//var str;
				//for(var i in jqXHR)
				//{
				//	 if(jqXHR.hasOwnProperty(i))
				//	{
				//		str += jqXHR[i];
				//	}
				//}
				//alert(str);
			<% end if %>
			document.location.reload();
			return false;
		}
	});
}

function goPage(page){
	<% if trim(research)="on" then %>
	    location.href='?page=' + page + '&research=on&menupos=<%=request("menupos")%>&deal=<%=deal%>&open=<%=open%>&keyword=<%=keyword%>&schWord=<%=schWord%>&state=<%=state%>';
	<% else %>
	    location.href="?page=" + page;
	<% end if %>
}

function goSearchPiece()
{
//	if ($("#deal").val()=="0"&&$("#open").val()=="A"&&$("#schWord").val()=="")
//	{
//		alert("�˻��� �ϱ� ���ؼ� ����, ���⿩��, Ű����˻� �� �ϳ���\n�������ֽðų� �Է����ּž� �մϴ�.");
//		return;
//	}
	document.frm1.submit();
}

function fnPieceDelact(idx, gubun)
{
	$("#frmDelidx").val(idx);

	var result

	if (gubun=="1")
	{
		result = confirm("������ �����Ͻðڽ��ϱ�?");
	}
	if (gubun=="2")
	{
		result = confirm("���̸� �����Ͻðڽ��ϱ�?");
	}
	if (gubun=="3")
	{
		result = confirm("����ƮŰ���带 �����Ͻðڽ��ϱ�?");
	}
	if (gubun=="4")
	{
		result = confirm("��ʸ� �����Ͻðڽ��ϱ�?");
	}
	if (gubun=="5")
	{
		result = confirm("ȸ�������� �����Ͻðڽ��ϱ�?");
	}

	if (result)
	{
		$.ajax({
			type:"GET",
			url:"/admin/sitemaster/piece/act_pieceDelete.asp",
			data:$("#frmpieceDel").serialize(),
			async:false,
			cache:true,
			success : function(Data, textStatus, jqXHR){
				if (jqXHR.readyState == 4) {
					if (jqXHR.status == 200) {
						if(Data!="") {
							res = Data.split("||");
							if (res[0]=="OK")
							{
								document.location.reload();
							}
							else
							{
								errorMsg = res[1].replace(">?n", "\n");
								alert(errorMsg);
								document.location.reload();
								return false;
							}
						} else {
							alert("�߸��� ���� �Դϴ�.");
							document.location.reload();
							return false;
						}
					}
				}
			},
			error:function(jqXHR, textStatus, errorThrown){
				alert("�߸��� ���� �Դϴ�.");
				<% if false then %>
					//var str;
					//for(var i in jqXHR)
					//{
					//	 if(jqXHR.hasOwnProperty(i))
					//	{
					//		str += jqXHR[i];
					//	}
					//}
					//alert(str);
				<% end if %>
				document.location.reload();
				return false;
			}
		});
	}
	else
	{
		return;
	}
}

</script>
<div class="">
	<div class="tab" style="margin:0 0 0 -1px;">
		<ul>
			<li class="col11 selected"><a href="#unitType01">��������</a></li>
			<li class="col11 "><a href="#unitType02">������ ����</a></li>
		</ul>
		<div class="managerInfo">
			<p><%=oPieceUser.FoneUser.Foccupation%> <strong><%=oPieceUser.FoneUser.Fnickname%></strong> <button type="button" class="memEdit">����</button></p>
			<p style="min-width:80px;">���� ���� <strong><%=pieceMyCnt(loginUserId)%></strong></p>
		</div>
	</div>

	<%' ��� �˻��� ���� %>
	<form name="frm1" id="frm1" method="get" action="/admin/sitemaster/piece/piece_terminal.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<%' search %>
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">���� :</label>
					<select class="formSlt" id="deal" name="deal" title="�ɼ� ����">
						<option value="0" <% If deal = "" Or deal = "0" Then %> selected <% End If %>>��ü</option>
						<option value="1" <% If deal = "1" Then %> selected <% End If %>>����</option>
						<option value="2" <% If deal = "2" Then %> selected <% End If %>>����</option>
						<option value="3" <% If deal = "3" Then %> selected <% End If %>>����Ʈ Ű����</option>
						<option value="4" <% If deal = "4" Then %> selected <% End If %>>���</option>
						<option value="5" <% If deal = "5" Then %> selected <% End If %>>ȸ�� ����</option>
					</select>
				</li>
				<li>
					<p class="formTit">���⿩�� :</p>
					<select class="formSlt" id="open" name="open" title="�ɼ� ����">
						<option value="A" <% If open = "" Or open = "A" Then %> selected <% End If %>>��ü</option>
						<option value="Y" <% If open = "Y" Then %> selected <% End If %>>����</option>
						<option value="N" <% If open = "N" Then %> selected <% End If %>>�����</option>
					</select>
				</li>
				<li>
					<p class="formTit">�������</p>
					<select class="formSlt" id="state" name="state" title="�ɼ� ����">
						<option value="" <% If state = ""  Then %> selected <% End If %>>��ü</option>
						<option value="1" <% If state = "1" Then %> selected <% End If %>>��ϴ��</option>
						<option value="2" <% If state = "2" Then %> selected <% End If %>>�̹��� ��Ͽ�û</option>
						<option value="3" <% If state = "3" Then %> selected <% End If %>>������ �۾���</option>
						<option value="4" <% If state = "4" Then %> selected <% End If %>>���¿�û</option>
						<option value="7" <% If state = "7" Then %> selected <% End If %>>����</option>
						<option value="8" <% If state = "8" Then %> selected <% End If %>>����</option>
						<option value="9" <% If state = "9" Then %> selected <% End If %>>����</option>
					</select>
				</li>
				<li>
					<p class="formTit">������</p>
					<input type="text" id="prevDate" name="prevDate" value="<%=schdate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "prevDate", trigger    : "prevDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">Ű���� �˻� :</label>
					<select class="formSlt" id="keyword" name="keyword" title="Ű���� �˻�">
						<option value="snum" <% If keyword = "snum" Then %> selected <% End If %>>��ȣ</option>
						<option value="stitle" <% If keyword ="" Or keyword = "stitle" Then %>selected<% End If %>>��������</option>
						<option value="sname" <% If keyword = "sname" Then %> selected <% End If %>>�ۼ���</option>
					</select>
					<input type="text" class="formTxt" id="schWord" name="schWord" style="width:400px" placeholder="Ű���带 �Է��Ͽ� �˻��ϼ���." value="<%=schWord%>" />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="�˻�" onclick="goSearchPiece();" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<!-- 20170824 ����
					<input type="button" class="btnOdrChg btn cBl1 fs12" value="��������" />
					-->
					<input type="button" class="btnRegist btn bold fs12" value="���" onclick="window.open('/admin/sitemaster/piece/popManagePiece.asp',null,'height=800,width=750,status=yes,toolbar=no,menubar=no,location=no');return false;" />
					<% If Trim(research)="on" Then %>
						<input type="button" class="btnRegist btn bold fs12" value="�˻��ʱ�ȭ" onclick="document.location.href='/admin/sitemaster/piece/piece_terminal.asp';" />
					<% End If %>
				</div>
				<!-- 20170824 ����
				<div class="ftLt">
					<p class="infoTxt">
						<span><img src='/images/ico_odrchg.png' alt='��������' /> �� ��� ���� ��, �Ʒ��� �̵� �� ����Ϸ� ��ư�� Ŭ�����ּ���.</span>
						!-- for dev msg:�˻����� ���� �� �������� ��ư Ŭ���� ����˴ϴ�. <span>�˻����� ���� �� ������ ������ �� �����ϴ�. <button type="button">�˻� �ʱ�ȭ</button></span> --
					</p>
				</div>
				-->
			</div>

			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">�� ��ϼ� : <strong><%=FormatNumber(oPieceList.FtotalCount, 0)%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
							<p style="width:80px">��ȣ(idx)</p>
							<p style="width:100px">����</p>
							<p class="">��������</p>
							<p style="width:50px">����</p>
							<p style="width:150px">�ۼ���<br/><span class="cRd1">����������</span></p>
							<p style="width:65px">���⿩��</p>
							<p style="width:120px">�����</p>
							<p style="width:120px">����������</p>
							<p style="width:120px">������</p>
							<p style="width:65px">�������</p>
							<p style="width:120px">����/����</p>
						</li>
					</ul>
					<ul id="sortable" class=" tbDataList">
						<% If Not(Trim(research)="on") Then %>
							<%'// ������ �����͸� ���� �ҷ��´�. %>
							<% If oPieceOpening.FResultCount > 0 Then %>
								<%'' for dev msg : ���������� ���õ� �׸��� li�� class="ui-state-disabled" �������ּ��� %>
								<li class="ui-state-disabled">
									<p style="width:80px">����<br/><a href="" onclick="window.open('http://m.10x10.co.kr/piece/piece_preview.asp?idx=<%=oPieceOpening.FOneOpening.FIdx%>',null,'height=720,width=375,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes');return false;" class="cBl1 tLine">[�̸�����]</a></p>
									<p style="width:100px">������</p>
									<p class="lt">
										<% If oPieceOpening.FOneOpening.Fgubun="1" Then %>
											<%=chrbyte(oPieceOpening.FOneOpening.Flisttext,75,"Y")%>
										<% Else %>
											<%=oPieceOpening.FOneOpening.Flisttitle%>
										<% End If %>
									</p>
									<p style="width:50px"><%=FormatNumber(oPieceOpening.FOneOpening.Fsnsbtncnt, 0)%></p>
									<p style="width:150px"><%=oPieceOpening.FOneOpening.Foccupation&" "&oPieceOpening.FOneOpening.Fnickname%><br/><span class="cRd1"><%=oPieceOpening.FOneOpening.Flastoccupation&" "&oPieceOpening.FOneOpening.Flastnickname%></span></p>
									<p style="width:65px">
										<% If oPieceOpening.FOneOpening.Fisusing="Y" Then %>
											����
										<% Else %>
											�����
										<% End If %>
									</p>
									<p style="width:120px"><%=Mid(Trim(oPieceOpening.FOneOpening.Fregdate), 1, 10)%><br/><%=Mid(Trim(oPieceOpening.FOneOpening.Fregdate), 11, 30)%></p>
									<p style="width:120px" class="cRd1"><%=Mid(Trim(oPieceOpening.FOneOpening.Flastupdate), 1, 10)%><br/><%=Mid(Trim(oPieceOpening.FOneOpening.Flastupdate), 11, 30)%></p>
									<p style="width:120px"><%=Mid(Trim(oPieceOpening.FOneOpening.Fstartdate), 1, 10)%><br/><%=Mid(Trim(oPieceOpening.FOneOpening.Fstartdate), 11, 30)%></p>
									<p style="width:65px;"><%=nowstatus(oPieceOpening.FOneOpening.Fstate)%></p>
									<p style="width:120px">
										<a href="" onclick="window.open('/admin/sitemaster/piece/popManagePiece.asp?idx=<%=oPieceOpening.FOneOpening.FIdx%>&page=<%=currpage%>&SearchDeal=<%=deal%>&SearchOpen=<%=open%>&SearchState=<%=state%>',null,'height=900,width=750,status=yes,toolbar=no,menubar=no,location=no');return false;" class="cBl1 tLine">[����]</a>
										<a href="" onclick="fnPieceDelact('<%=oPieceOpening.FOneOpening.FIdx%>', '<%=Trim(oPieceOpening.FOneOpening.Fgubun)%>');return false;" class="cBl1 tLine">[����]</a>
									</p>
								</li>
							<% End If %>
						<% End If %>

						<%'// ������ �����͸� ������ ����Ʈ�� �����´�. %>
						<% If oPieceList.FResultcount > 0 Then %>
							<% For i=0 To oPieceList.Fresultcount-1 %>
							<li>
								<!--p style="width:80px"><%=(oPieceList.FtotalCount - pagesize * (currpage-1) - i)%></p-->
								<p style="width:80px"><%=oPieceList.FPieceList(i).FIdx%><br/><a href="" onclick="window.open('http://m.10x10.co.kr/piece/piece_preview.asp?idx=<%=oPieceList.FPieceList(i).FIdx%>',null,'height=720,width=375,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes');return false;" class="cBl1 tLine">[�̸�����]</a></p>
								<p style="width:100px">
									<% Select Case Trim(oPieceList.FPieceList(i).Fgubun) %>
										<% Case "1" %>
											����
										<% Case "2" %>
											����
										<% Case "3" %>
											����ƮŰ����
										<% Case "4" %>
											���
										<% Case "5" %>
											ȸ������
									<% End Select %>
								</p>
								<p class="lt">
									<% If oPieceList.FPieceList(i).Fgubun="1" Then %>
										<%=chrbyte(oPieceList.FPieceList(i).Flisttext,75,"Y")%>
									<% Else %>
										<%=oPieceList.FPieceList(i).Flisttitle%>
									<% End If %>
								</p>
								<p style="width:50px"><%=oPieceList.FPieceList(i).Fsnsbtncnt%></p>
								<p style="width:150px"><%=oPieceList.FPieceList(i).Foccupation&" "&oPieceList.FPieceList(i).Fnickname%><br/><span class="cRd1"><%=oPieceList.FPieceList(i).Flastoccupation&" "&oPieceList.FPieceList(i).Flastnickname%></span></p>
								<p style="width:65px">
								<% If Trim(oPieceList.FPieceList(i).Fisusing)="Y" Then %>
									����
								<% Else %>
									�����
								<% End If %>
								</p>
								<p style="width:120px"><%=Mid(Trim(oPieceList.FPieceList(i).Fregdate), 1, 10)%><br/><%=Mid(Trim(oPieceList.FPieceList(i).Fregdate), 11, 30)%></p>
								<p style="width:120px" class="cRd1"><% If oPieceList.FPieceList(i).Flastnickname <> "" Then %><%=Mid(Trim(oPieceList.FPieceList(i).Flastupdate), 1, 10)%><br/><%=Mid(Trim(oPieceList.FPieceList(i).Flastupdate), 11, 30)%><% End If %></p>
								<p style="width:120px"><%=Mid(Trim(oPieceList.FPieceList(i).Fstartdate), 1, 10)%><br/><%=Mid(Trim(oPieceList.FPieceList(i).Fstartdate), 11, 30)%></p>
								<p style="width:65px"><%=nowstatus(oPieceList.FPieceList(i).Fstate)%></p>
								<p style="width:120px">
									<a href="" onclick="window.open('/admin/sitemaster/piece/popManagePiece.asp?idx=<%=oPieceList.FPieceList(i).FIdx%>&SearchDeal=<%=deal%>&SearchOpen=<%=open%>&SearchState=<%=state%>',null,'height=900,width=750,status=yes,toolbar=no,menubar=no,location=no');return false;" class="cBl1 tLine">[����]</a>
									<a href="" onclick="fnPieceDelact('<%=oPieceList.FPieceList(i).FIdx%>', '<%=Trim(oPieceList.FPieceList(i).Fgubun)%>');return false;" class="cBl1 tLine">[����]</a>
								</p>
							</li>
							<% Next %>
						<% End If %>
					</ul>
					<div class="ct tPad20 cBk1">
						<%= fnDisplayPaging_New2017(currpage, oPieceList.FtotalCount, pagesize, 10, "goPage") %>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<div class="lyrBox">
	<div class="pieceMember">
		<strong>����� � ����ΰ���?</strong>
		<p class="tPad10">����� ������ Piece���� ����� ������ �Է����ּ���.</p>
		<div class="whoAreYou">
			<p class="ftLt">
				<select class="formSlt" style="width:100px; height:30px;" name="selectOccupation" id="selectOccupation">
					<option value="A">��������</option>
					<option value="Member" <% If oPieceUser.FoneUser.Foccupation="Member" Then %>selected<% End If %>>Member</option>
					<option value="Planner" <% If oPieceUser.FoneUser.Foccupation="Planner" Then %>selected<% End If %>>Planner</option>
					<option value="Designer" <% If oPieceUser.FoneUser.Foccupation="Designer" Then %>selected<% End If %>>Designer</option>
					<option value="Publisher" <% If oPieceUser.FoneUser.Foccupation="Publisher" Then %>selected<% End If %>>Publisher</option>
					<option value="Developer" <% If oPieceUser.FoneUser.Foccupation="Developer" Then %>selected<% End If %>>Developer</option>
					<option value="MD" <% If oPieceUser.FoneUser.Foccupation="MD" Then %>selected<% End If %>>MD</option>
					<option value="Editor" <% If oPieceUser.FoneUser.Foccupation="Editor" Then %>selected<% End If %>>Editor</option>
				</select>
			</p>
			<p class="ftRt">
				<input type="text" placeholder="�г���" class="formTxt" style="height:30px;" name="usernickname" id="usernickname" value="<%=oPieceUser.FoneUser.Fnickname%>" />
			</p>
		</div>
		<p>
			<input type="button" value="Ȯ��" class="cRd1" style="width:100px; height:30px;" onclick="fnPieceUseract();return false;" />
		</p>
	</div>
</div>
<form name="frmpieceUser" id="frmpieceUser">
	<input type="hidden" name="frmmode" id="frmmode">
	<input type="hidden" name="frmoccupation" id="frmoccupation">
	<input type="hidden" name="frmnickname" id="frmnickname">
	<input type="hidden" name="frmadminid" id="frmadminid" value="<%=loginUserId%>">
	<input type="hidden" name="frmidx" id="frmidx" value="<%=oPieceUser.FoneUser.Fidx%>">
</form>
<form name="frmpieceDel" id="frmpieceDel">
	<input type="hidden" name="frmDeladminid" id="frmDeladminid" value="<%=loginUserId%>">
	<input type="hidden" name="frmDelidx" id="frmDelidx">
</form>
<div class="dimmed"></div>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function() {
	$(".btnOdrChg").on('click',function() {
		if ($("#sortable").hasClass('sortable')) {
			$("#sortable").removeClass('sortable');
			$("#sortable li p:first-child").html("901"); //����Ʈ index�� ���Բ�
			$("#sortable li.ui-state-disabled p:first-child").html("����");
			$("#sortable").sortable("destroy");
			$(".btnOdrChg").attr("value", "��������");
			//$(".btnOdrChg").prop("disabled", true); //�˻����� ����� �������� ��ư ��Ȱ��ȭ
			$(".btnRegist").prop("disabled", false);
			$(".infoTxt").hide();
		} else {
			$("#sortable").addClass('sortable');
			$("#sortable li p:first-child").html("<img src='/images/ico_odrchg.png' alt='��������' />");
			$("#sortable li.ui-state-disabled p:first-child").html("����");
			$("#sortable").sortable({
				placeholder:"handling",
				items:"li:not(.ui-state-disabled)"
			}).disableSelection();
			$(".btnOdrChg").attr("value", "����Ϸ�");
			//$(".btnOdrChg").prop("disabled", false);
			$(".btnRegist").prop("disabled", true);
			$(".infoTxt").show();
		}
	});

	$(".memEdit").on('click',function() {
		$(".dimmed").show();
		$(".lyrBox").show();
	});

	<% if oPieceUser.FResultCount < 1 then %>
		$(".dimmed").show();
		$(".lyrBox").show();
	<% end if %>

});
</script>

</body>
</html>
<%
	Set oPieceUser = Nothing
	Set oPieceList = Nothing
	Set oPieceOpening = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
