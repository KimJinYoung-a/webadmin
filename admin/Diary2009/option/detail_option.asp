<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.09 �ѿ�� ����
'	Description : ���̾���丮
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/DiaryCls.asp"-->
<%
Dim i ,a ,b ,oip ,oip_keyword , oip_contents ,DiaryID ,inttr
	DiaryID = request("DiaryID")
	inttr=0	
%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript">

	function viewplay(idx){
		frm.idx.value = idx;
		frm.submit();
	}
	
	function getsubmit(){
		frm_edit.mode.value = 'edit';	
		frm_edit.mode_type.value = 'keyword';
		frm_edit.submit();
	}
	
	function new_submit(){	
		var new_submit;
		new_submit = window.open("/admin/diary2009/option/keyword_option_new.asp", "new_submit","width=1024,height=200,scrollbars=yes,resizable=yes");
		new_submit.focus();
	}
	
	function keyword_change(DiaryID,option_value,boxvalue){
		if (boxvalue == '0') {
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'keyword';
		frm_temp.mode_type.value = 'insert';		
		frm_temp.action = "/admin/diary2009/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}else{
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'keyword';
		frm_temp.mode_type.value = 'delete';		
		frm_temp.action = "/admin/diary2009/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}
	}
	
	function contents_change(DiaryID,option_value,boxvalue){
		if (boxvalue == '0') {
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'contents';
		frm_temp.mode_type.value = 'insert';		
		frm_temp.action = "/admin/diary2009/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}else{
		frm_temp.keyword_option.value =  option_value;
		frm_temp.mode.value = 'contents';
		frm_temp.mode_type.value = 'delete';		
		frm_temp.action = "/admin/diary2009/option/detail_option_process.asp";
		frm_temp.target = 'view';
		frm_temp.submit();
		}
	}	
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<span>
		�� �����Ͻø� �ٷ� �Ǽ����� ���� �˴ϴ�. �ʹ� ����..������ ������!<br>
		�ɼ����� ���� = Y , �ɼ��� �ƴҰ�� = N
		</span>

		<!-- Ű���� ����-->
		<div class="tPad15">
		<form name="frm_keyword" action="" method="post">
		<input type="hidden" name="mode" >
		<input type="hidden" name="DiaryID" value="<%=DiaryID%>">
		<table class="tbType1 listTb">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td align="center" colspan=5>Ű����</td>
			</tr>
			<tr align="center" bgcolor="#FFFFFF">
				<% 
				'//Ű���� ���� ��������
				set oip = new DiaryCls
					oip.fkeyword_type()

				for i = 0 to oip.FResultCount -1
				if oip.FItemList(i).ftype <> "" then 
				%>	
				<td style="vertical-align:top;">
					<table class="tbType1 listTb">
						<tr align="center" bgcolor="#FFFFFF">
							<td align="conter">
								<font color="blue"><%= oip.FItemList(i).ftype %></font>	
							</td>
						</tr>
						<%
						'// ����Ű���� �Ѹ��� 
						set oip_keyword = new DiaryCls
						oip_keyword.frecttype = oip.FItemList(i).ftype
						oip_keyword.frectdiaryid = DiaryID
						oip_keyword.fkeyword_option_value()
						
						if oip_keyword.FResultCount <> 0 then

						for a = 0 to oip_keyword.FResultCount -1
							
						%>
						<tr>
							<td align="left">
								<%= oip_keyword.FItemList(a).foption_value %>	<%= oip_keyword.FItemList(a).fcontents_idx %>
								<%
									If oip.FItemList(i).ftype = "color" Then
										Response.Write " <img src='http://fiximage.10x10.co.kr/web2011/diarystory2012/search_" & replace(oip_keyword.FItemList(a).foption_value,"Lightgray","gray") & ".gif'>"
									End If
								%>
								<select name="<%= oip_keyword.FItemList(a).foption_value %>"  onchange="keyword_change('DiaryID','<%= oip_keyword.FItemList(a).fidx %>',this.value);">
									<option value="1">N</option>
									<option value="0" <% if oip_keyword.FItemList(a).fkeyword_option_count <>"" then response.write " selected" %>>Y</option>							
								</select>				
							</td>
						</tr>
						<% 
						next
								
						end if %>
					</table>
				</td>
				<% 
					end if 
					next 
				%>
			</tr>   
		</table>
		</form>		
		<!-- Ű���� ��-->
		<form name="frm_temp" action="" method="post">
			<input type="hidden" name="mode">
			<input type="hidden" name="mode_type" >
			<input type="hidden" name="keyword_option">	
			<input type="hidden" name="DiaryID" value="<%=DiaryID%>">
		</form>
		<iframe frameborder=0 name="view" id="view" width="0" height="0"></iframe>
		</div>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->