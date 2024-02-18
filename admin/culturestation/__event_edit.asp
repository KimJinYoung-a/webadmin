<%@ language="VBScript" %>
<% option Explicit %>
<%
'###########################################################
' Description : Culture Station Event 수정  
' History : 2008.04.02 한용민 생성
'           2012.01.12 허진원; 모바일 항목 추가
'			2015.01.23 유태욱 진행업체 추가(evt_partner)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->
<link rel="stylesheet" href="/admin/culturestation/daumeditor/css/editor.css" type="text/css"/>
<script src="/admin/culturestation/daumeditor/js/editor_loader.js?environment=" type="text/javascript"></script>
<%
dim m_evtbn_code
dim YearUse ,startdate,enddate,evt_code ,evt_name , isusing , ticket_isusing , comment, write_work
dim eventdate , listimgName , barnerimgName , main1imgName , main2imgName , image_main_link, barner2imgName, barner3imgName
dim main3imgName, main4imgName, main5imgName , evt_type, evt_comment, evt_partner
dim m_isusing, m_img_icon, m_img_main1, m_img_main2, m_main_content, m_cmt_desc, m_sortNo
Dim edid , emid, evt_kind

	YearUse = 2009
	evt_code = requestCheckVar(request("evt_code"),10)
	evt_kind = requestCheckVar(request("evt_kind"),1)
	isusing = "N"
	m_isusing = "N"
	m_sortNo = 1

dim oip, i
set oip = new cevent_list
	oip.frectevt_code = evt_code
	if evt_code <> "" then
		oip.fevent_list()
		evt_type = oip.FItemList(0).fevt_type
		evt_code = oip.FItemList(0).fevt_code
		evt_name = oip.FItemList(0).fevt_name
		evt_comment = oip.FItemList(0).fevt_comment
		
		evt_partner = oip.FItemList(0).fevt_partner
		
		isusing = oip.FItemList(0).fisusing
		ticket_isusing = oip.FItemList(0).fticket_isusing
		comment = oip.FItemList(0).fcomment
		startdate = oip.FItemList(0).fstartdate
		enddate = oip.FItemList(0).fenddate
		eventdate = oip.FItemList(0).feventdate
		listimgName = oip.FItemList(0).fimage_list
		barner2imgName = oip.FItemList(0).fimage_barner2		'2013년 추가
		barner3imgName = oip.FItemList(0).fimage_barner3		'2013년 추가
		main1imgName = oip.FItemList(0).fimage_main
		main2imgName = oip.FItemList(0).fimage_main2
		main3imgName = oip.FItemList(0).fimage_main3
		main4imgName = oip.FItemList(0).fimage_main4
		main5imgName = oip.FItemList(0).fimage_main5
		image_main_link = oip.FItemList(0).fimage_main_link
		write_work = oip.FItemList(0).fwrite_work

		edid = oip.FItemList(0).fedid
		emid = oip.FItemList(0).femid
		evt_kind = oip.FItemList(0).fevt_kind
		m_isusing		= oip.FItemList(0).fm_isusing
		m_img_icon		= oip.FItemList(0).fm_img_icon
		m_img_main1		= oip.FItemList(0).fm_img_main1
		m_img_main2		= oip.FItemList(0).fm_img_main2
		m_main_content	= nl2blank(oip.FItemList(0).fm_main_content)
		m_cmt_desc		= nl2blank(oip.FItemList(0).fm_cmt_desc)
		m_sortNo		= oip.FItemList(0).fm_sortNo
		
		m_evtbn_code	= oip.FItemList(0).fm_evtbn_code

	end If

	If IsNull(edid) Then edid = " "
	If IsNull(emid) Then emid = " "
%>
<!-- daumeditor head ------------------------->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=10" /> 
<link rel="stylesheet" href="/admin/culturestation/daumeditor/css/editor.css" type="text/css" charset="euc-kr"/>
<script src="/admin/culturestation/daumeditor/js/editor_loader.js" type="text/javascript" charset="euc-kr"></script>
<script src="/admin/culturestation/daumeditor/js/editor_creator.js" type="text/javascript" charset="euc-kr"></script>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<!-- //daumeditor head ------------------------->
<script language="javascript">

	function image_view(src){
		var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
		image_view.focus();
	}

	//''이미지 저장
	function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,fheight,thumb){
		var popImg = window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
		document.imginputfrm.divName.value=divnm;
		document.imginputfrm.inputname.value=iptNm;
		document.imginputfrm.ImagePath.value = vPath;
		document.imginputfrm.maxFileSize.value = Fsize;
		document.imginputfrm.maxFileWidth.value = Fwidth;
		document.imginputfrm.maxFileheight.value = fheight;	
		document.imginputfrm.makeThumbYn.value = thumb;
		document.imginputfrm.orgImgName.value = document.getElementById(iptNm).value;
		document.imginputfrm.target='imginput';
		document.imginputfrm.action='event_img_input.asp';
		document.imginputfrm.submit();
		popImg.focus();
	}

	document.domain = "10x10.co.kr";

	//''이벤트저장
	function event_reg(mode){
		
		if(document.frm.evt_type.value==''){
			alert('이벤트 구분을 입력하셔야 합니다.');
			return false;
		}

		var content = Editor.getContent();
		//alert(content);
		Editor.switchEditor("2");
		var content2 = Editor.getContent();
		//alert(content2);
		document.getElementById("m_cmt_desc").value = content;
		document.getElementById("m_main_content").value = content2;

		if (mode == 'add'){
			frm.mode.value='add';
		}else if(mode == 'edit'){
			frm.mode.value='edit';
		}

		frm.submit();		
	}
	
	function delimage(gubun)
	{
		var aa = eval("document.frm."+gubun+"");
		aa.value = "";
		frm.mode.value = "edit";
		frm.isimgdel.value = "o";
		frm.submit();
	}
	function write_event(){
		if(document.getElementById('hand').checked){
			document.getElementById('main1').style.display = "none";
			document.getElementById('main2').style.display = "none";
			document.getElementById('main3').style.display = "none";
			document.getElementById('main4').style.display = "none";
			document.getElementById('main5').style.display = "none";
			document.getElementById('main6').style.display = "block";
		}else{
			document.getElementById('main1').style.display = "block";
			document.getElementById('main2').style.display = "block";
			document.getElementById('main3').style.display = "block";
			document.getElementById('main4').style.display = "block";
			document.getElementById('main5').style.display = "block";
			document.getElementById('main6').style.display = "none";
		}
	}
	function jsManageEventImageNew(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/cultureManageDir.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}

	//모바일 사용여부 변경
	function chgMobileUsing(sel) {
		if(sel=="Y") {
			for(var i=0;i<9;i++) {
				document.getElementById("trMobile"+i).style.display="";
			}
			document.getElementById("lyMobileSort").style.display="";
		} else {
			for(var i=0;i<9;i++) {
				document.getElementById("trMobile"+i).style.display="none";
			}
			document.getElementById("lyMobileSort").style.display="none";
		}
	}
	

	$(document).ready(function(){
		$('#evt_type').change(function(){
			if($('#evt_type').val()=="0"){
				$("#evt_kind option").remove();
				$("#evt_kind").append("<option value='1'>영화</option>");
				$("#evt_kind").append("<option value='2'>연극</option>");
				$("#evt_kind").append("<option value='3'>공연</option>");
				$("#evt_kind").append("<option value='4'>뮤지컬</option>");
				
			}else{
				$("#evt_kind option").remove();
				$("#evt_kind").append("<option value='5'>도서</option>");
			}
		});
	});
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="frm" method="post" action="/admin/culturestation/event_edit_process.asp">
	<input type="hidden" name="mode" >
	<input type="hidden" name="isimgdel" >
	<input type="hidden" name="evt_code" value="<%=evt_code%>">
	<input type="hidden" name="m_main_content" id="m_main_content" value="<%=m_main_content%>">
	<input type="hidden" name="m_cmt_desc" id="m_cmt_desc" value="<%=m_cmt_desc%>">
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>Event 구분</b><br></td>
		<td><select name="evt_type" id="evt_type" class="select">
				<option value="">선택</option>
				<option value="0" <% if evt_type = "0" then response.write " selected" %>>느껴봐</option>
				<option value="1" <% if evt_type = "1" then response.write " selected" %>>읽어봐</option>
				<!-- <option value="2" <% if evt_type = "2" then response.write " selected" %>>들어봐</option> -->
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>Event 종류</b><br></td>
		<td><select name="evt_kind" id="evt_kind" class="select">
				<option value="0" <% if evt_kind = "0" then response.write " selected" %>>영화</option>
				<option value="1" <% if evt_kind = "1" then response.write " selected" %>>연극</option>
				<option value="2" <% if evt_kind = "2" then response.write " selected" %>>공연</option>
				<option value="3" <% if evt_kind = "3" then response.write " selected" %>>뮤지컬</option>
				<option value="4" <% if evt_kind = "4" then response.write " selected" %>>도서</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>Event명</b><br></td>
		<td><input type="text" name="evt_name" size="50" value="<%= evt_name %>" class="text">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>진행업체</b><br></td>
		<td><input type="text" name="evt_partner" size="32" maxlength="32" value="<%= evt_partner %>" class="text">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>추가코멘트</b><br></td>
		<td><input type="text" name="evt_comment" size="32" maxlength="32" value="<%= evt_comment %>" class="text"> ※사은품 등의 추가 정보
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>사용여부</b><br></td>
		<td>
			Web : <input type="radio" name="isusing" id="isusing" value="Y"<% if isusing = "Y" then response.write " checked" %>>사용
			<input type="radio" name="isusing" id="isusing" value="N"<% if isusing = "N" then response.write " checked" %>>사용안함&nbsp;&nbsp;
			Mobile : <input type="radio" name="m_isusing" id="m_isusing" value="Y"<% if m_isusing = "Y" then response.write " checked" %> onclick="chgMobileUsing(this.value)">사용
			<input type="radio" name="m_isusing" id="m_isusing" value="N"<% if m_isusing = "N" then response.write " checked" %> onclick="chgMobileUsing(this.value)">사용안함
			<span id="lyMobileSort" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
			[목록순서
			<input type="text" name="m_sortNo" value="<%=m_sortNo%>" class="text" style="width:24px; text-align:center;">
			]</span>
		</td>
	</tr>
	<!--<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>티켓양도사용여부</b><br></td>
		<td><select name="ticket_isusing">
				<option value="Y" <% if ticket_isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if ticket_isusing = "N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>-->	
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>기능</b><br></td>
		<td><input type="checkbox" name="comment" value="ON" <% if comment = "ON" then response.write " checked" %>>코맨트사용
			<input type="checkbox" id = "hand" name="write_work" value="Y" <% if write_work = "Y" then response.write " checked" %> onclick="write_event()">수작업사용
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>Event 기간</b><br></td>
		<td>
			<input type="text" name="startdate" size=10 value="<%= startdate %>" class="text">			
			<a href="javascript:calendarOpen3(frm.startdate,'시작일',frm.startdate.value)">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
			<input type="text" name="enddate" size=10  value="<%= left(enddate,10) %>" class="text">
			<a href="javascript:calendarOpen3(frm.enddate,'마지막일',frm.enddate.value)">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
		</td>
	</tr>			
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>Event 당첨일</b><br></td>
		<td>
			<input type="text" name="eventdate" size=10 value="<%= eventdate %>" class="text">			
			<a href="javascript:calendarOpen3(frm.eventdate,'시작일',frm.eventdate.value)">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>WD담당</b><br></td>
		<td>
			<%sbGetDesignerid "selDId",edid,""%>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>마케팅담당</b><br></td>
		<td>
			<%sbGetMKTid2 "selMId",emid,""%>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>기본 이미지</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('listdiv','listimgName','list','750 ','1110','5000','true');"  class="button"> 
			(<b><font color="red">200x165</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="listimgName" id="listimgName" value="<%= listimgName %>">
			<div align="right" id="listdiv">
				<% if listimgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/list/<%= oip.FItemList(i).fimage_list %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/list/<%= oip.FItemList(i).fimage_list %>" width="25" height="25" border=0>
					</a>
				<% end if %>	
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>배너 이미지</b><br></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('barnerdiv','barnerimgName','barner','200','200','65','false');"  class="button"> 
			(<b><font color="red">200x65</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="barnerimgName" id="barnerimgName" value="<%= barnerimgName %>">
			<div align="right" id="barnerdiv">
				<% if barnerimgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/barner/<%= oip.FItemList(i).fimage_barner %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/barner/<%= oip.FItemList(i).fimage_barner %>" width="25" height="25"  border=0>
					</a>
				<% end if %>
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>배너(L) #2013</b><br></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('barner2div','barner2imgName','barner2','300','564','705','false');"  class="button"> 
			(<b><font color="red">564x705</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="barner2imgName" id="barner2imgName" value="<%= barner2imgName %>">
			<div align="right" id="barner2div">
				<% if barner2imgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/barner2/<%= oip.FItemList(i).fimage_barner2 %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/barner2/<%= oip.FItemList(i).fimage_barner2 %>" width="25" height="25"  border=0>
					</a>
				<% end if %>
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>배너(S) #2013</b><br></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('barner3div','barner3imgName','barner3','200','276','300','false');"  class="button"> 
			(<b><font color="red">276x300</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="barner3imgName" id="barner3imgName" value="<%= barner3imgName %>">
			<div align="right" id="barner3div">
				<% if barner3imgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/barner3/<%= oip.FItemList(i).fimage_barner3 %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/barner3/<%= oip.FItemList(i).fimage_barner3 %>" width="25" height="25"  border=0>
					</a>
				<% end if %>
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="main1" <%If write_work = "Y" Then response.write "style=display:none;" %>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>메인 이미지1</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main1div','main1imgName','main1','600','898','5000','false');"  class="button"> 
			(<b><font color="red">가로898</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="main1imgName" id="main1imgName" value="<%= main1imgName %>">
			<div align="right" id="main1div">
				<% if main1imgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/main1/<%= main1imgName %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/main1/<%= main1imgName %>" width="25" height="25"  border=0>
					</a>
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="main2" <%If write_work = "Y" Then response.write "style=display:none;" %>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>메인 이미지2</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main2div','main2imgName','main2','600','898','5000','false');"  class="button"> 
			(<b><font color="red">가로898</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main2imgName" id="main2imgName" value="<%= main2imgName %>">
			<div align="right" id="main2div">
				<% if main2imgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/main2/<%= main2imgName %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/main2/<%= main2imgName %>" width="25" height="25"  border=0>
					</a>
					<a href="javascript:delimage('main2imgName');">[삭제]</a>
				<% end if %>				
			</div>
		</td>
	</tr>
		
	<tr bgcolor="#FFFFFF" id="main3" <%If write_work = "Y" Then response.write "style=display:none;" %>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>메인 이미지3</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main3div','main3imgName','main3','600','898','5000','false');"  class="button"> 
			(<b><font color="red">가로898</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main3imgName" id="main3imgName" value="<%= main3imgName %>">
			<div align="right" id="main3div">
				<% if main3imgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/main3/<%= main3imgName %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/main3/<%= main3imgName %>" width="25" height="25"  border=0>
					</a>
					<a href="javascript:delimage('main3imgName');">[삭제]</a>
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="main4" <%If write_work = "Y" Then response.write "style=display:none;" %>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>메인 이미지4</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main4div','main4imgName','main4','600','898','5000','false');"  class="button"> 
			(<b><font color="red">가로898</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main4imgName" id="main4imgName" value="<%= main4imgName %>">
			<div align="right" id="main4div">
				<% if main4imgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/main4/<%= main4imgName %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/main4/<%= main4imgName %>" width="25" height="25"  border=0>
					</a>
					<a href="javascript:delimage('main4imgName');">[삭제]</a>
				<% end if %>				
			</div>
		</td>
	</tr>
	
	<tr bgcolor="#FFFFFF" id="main5" <%If write_work = "Y" Then response.write "style=display:none;" %>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>메인 이미지5</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('main5div','main5imgName','main5','600','898','5000','false');"  class="button"> 
			(<b><font color="red">가로898</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="main5imgName" id="main5imgName" value="<%= main5imgName %>">
			<div align="right" id="main5div">
				<% if main5imgName <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/culturestation/<%= yearUse %>/main5/<%= main5imgName %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl%>/culturestation/<%= yearUse %>/main5/<%= main5imgName %>" width="25" height="25"  border=0>
					</a>
					<a href="javascript:delimage('main5imgName');">[삭제]</a>
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="main6" <%If write_work = "Y" Then response.write "style=display:block;" else response.write "style=display:none;"%>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>수작업 이미지</b></td>
		<td>
			<input type="button" value="이미지관리"  onclick="jsManageEventImageNew('<%= evt_code%>')" class="input_b">
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2"><b>메인이미지 이미지맵 & 링크 코드</b> <font color="red"> map이름 절대 수정금지 !!</font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<textarea rows="15" name="image_main_link" style="width:100%" class="textarea"><%=chkIIF(evt_code<>"",image_main_link,"<map name=""ImgMap1""></map>" & vbCrLf & "<map name=""ImgMap2""></map>" & vbCrLf & "<map name=""ImgMap3""></map>" & vbCrLf & "<map name=""ImgMap4""></map>" & vbCrLf & "<map name=""ImgMap5""></map>" & vbCrLf) %></textarea>
		</td>
	</tr>	
<!----- ## 모바일 폼 시작 --------------------------------->
	<tr bgcolor="#F0F0FF" id="trMobile0" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" colspan="2"><b>모바일 사이트 내용</b></td>
	</tr>
	<tr bgcolor="#FFFFFF" id="trMobile1" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>모바일 목록</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('m_icondiv','m_img_icon','mIcon','200','180','210','false');"  class="button"> 
			(<b><font color="red">180x210</font></b><b><font color="red"></font></b>)
			<input type="hidden" name="m_img_icon" id="m_img_icon" value="<%= m_img_icon %>">
			<div align="right" id="m_icondiv">
				<% if m_img_icon <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl & "/culturestation/" & yearUse & "/mIcon/" & m_img_icon %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl & "/culturestation/" & yearUse & "/mIcon/" & m_img_icon %>" width="25" height="25" border=0></a>
				<% end if %>	
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="trMobile2" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>코멘트 안내 이미지</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('m_img1div','m_img_main1','mMain1','600','529','4500','false');"  class="button"> 
			(<b><font color="red">가로529</font></b>)
			<input type="hidden" name="m_img_main1" id="m_img_main1" value="<%= m_img_main1 %>">
			<div align="right" id="m_img1div">
				<% if m_img_main1 <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl & "/culturestation/" & yearUse & "/mMain1/" & m_img_main1 %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl & "/culturestation/" & yearUse & "/mMain1/" & m_img_main1 %>" width="25" height="25" border=0></a>
				<% end if %>				
			</div>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="trMobile3" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>모바일 이벤트 배너</b></td>
		<td>
			<input type="button" value="이미지 넣기" onclick="jsImgInput('m_img2div','m_img_main2','mMain2','600','640','4500','false');"  class="button"> 
			(<b><font color="red">가로529</font></b><b><font color="red"></font></b>) 없으면 등록하지 마세요.
			<input type="hidden" name="m_img_main2" id="m_img_main2" value="<%= m_img_main2 %>">
			<div align="right" id="m_img2div">
				<% if m_img_main2 <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl & "/culturestation/" & yearUse & "/mMain2/" & m_img_main2 %>');" onfocus="this.blur()">
					<img src="<%=webImgUrl & "/culturestation/" & yearUse & "/mMain2/" & m_img_main2 %>" width="25" height="25" border=0></a>
					<a href="javascript:delimage('m_img_main2');">[삭제]</a>
				<% end if %>				
			</div>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF" id="trMobile8"  <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" bgcolor="<%=adminColor("sky")%>"><b>이벤트 링크 코드</b><br></td>
		<td><input type="text" name="m_evtbn_code" size="20" value="<%= m_evtbn_code %>" class="text">
			&lt;&lt;이벤트 코드 입력
		</td>
	</tr>

	<tr bgcolor="#FFFFFF" id="trMobile4" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" colspan="2"><b>모바일 컨텐츠 정보입력</b> <!--<font color="red">&lt;P&gt;클래스 절대 수정금지!!</font>--></td>
	</tr>
	<tr bgcolor="#FFFFFF" id="trMobile5" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" colspan="2">
			<!-- #include virtual="/admin/culturestation/daumeditor/editor.asp"-->
		</td>
	</tr>
<script type="text/javascript">
	var config2 = {
		txHost: '', /* 런타임 시 리소스들을 로딩할 때 필요한 부분으로, 경로가 변경되면 이 부분 수정이 필요. ex) http://xxx.xxx.com */
		txPath: '', /* 런타임 시 리소스들을 로딩할 때 필요한 부분으로, 경로가 변경되면 이 부분 수정이 필요. ex) /xxx/xxx/ */
		txService: 'sample', /* 수정필요없음. */
		txProject: 'sample', /* 수정필요없음. 프로젝트가 여러개일 경우만 수정한다. */
		initializedId: "2", /* 대부분의 경우에 빈문자열 */
		wrapper: "tx_trex_container2", /* 에디터를 둘러싸고 있는 레이어 이름(에디터 컨테이너) */
		form: "frm"+"", /* 등록하기 위한 Form 이름 */
		txIconPath: "images/icon/editor/", /*에디터에 사용되는 이미지 디렉터리, 필요에 따라 수정한다. */
		txDecoPath: "images/deco/contents/", /*본문에 사용되는 이미지 디렉터리, 서비스에서 사용할 때는 완성된 컨텐츠로 배포되기 위해 절대경로로 수정한다. */
		canvas: {
			styles: {
				color: "#123456", /* 기본 글자색 */
				fontFamily: "굴림", /* 기본 글자체 */
				fontSize: "10pt", /* 기본 글자크기 */
				backgroundColor: "#fff", /*기본 배경색 */
				lineHeight: "1.5", /*기본 줄간격 */
				padding: "8px" /* 위지윅 영역의 여백 */
			},
			showGuideArea: false
		},
		events: {
			preventUnload: false
		},
		sidebar: {
			attachbox: {
				show: true
			},
			attacher: {
				file: {
					popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
				}
			}
		},
		size: {
			contentWidth: 700 /* 지정된 본문영역의 넓이가 있을 경우에 설정 */
		}
	};
</script>
	<tr bgcolor="#FFFFFF" id="trMobile6" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" colspan="2"><b>모바일 컨텐츠 설명입력</b> <!--<font color="red">&lt;P&gt;클래스 절대 수정금지!!</font>--></td>
	</tr>

	<tr bgcolor="#FFFFFF" id="trMobile7" <%=chkIIF(m_isusing="Y","","style='display:none;'")%>>
		<td align="center" colspan="2">
			<!-- #include virtual="/admin/culturestation/daumeditor/editor2.asp"-->
		</td>
	</tr>
<script type="text/javascript">
	var config3 = {
		txHost: '', /* 런타임 시 리소스들을 로딩할 때 필요한 부분으로, 경로가 변경되면 이 부분 수정이 필요. ex) http://xxx.xxx.com */
		txPath: '', /* 런타임 시 리소스들을 로딩할 때 필요한 부분으로, 경로가 변경되면 이 부분 수정이 필요. ex) /xxx/xxx/ */
		txService: 'sample', /* 수정필요없음. */
		txProject: 'sample', /* 수정필요없음. 프로젝트가 여러개일 경우만 수정한다. */
		initializedId: "3", /* 대부분의 경우에 빈문자열 */
		wrapper: "tx_trex_container3", /* 에디터를 둘러싸고 있는 레이어 이름(에디터 컨테이너) */
		form: "frm"+"", /* 등록하기 위한 Form 이름 */
		txIconPath: "images/icon/editor/", /*에디터에 사용되는 이미지 디렉터리, 필요에 따라 수정한다. */
		txDecoPath: "images/deco/contents/", /*본문에 사용되는 이미지 디렉터리, 서비스에서 사용할 때는 완성된 컨텐츠로 배포되기 위해 절대경로로 수정한다. */
		canvas: {
			styles: {
				color: "#123456", /* 기본 글자색 */
				fontFamily: "굴림", /* 기본 글자체 */
				fontSize: "10pt", /* 기본 글자크기 */
				backgroundColor: "#fff", /*기본 배경색 */
				lineHeight: "1.5", /*기본 줄간격 */
				padding: "8px" /* 위지윅 영역의 여백 */
			},
			showGuideArea: false
		},
		events: {
			preventUnload: false
		},
		sidebar: {
			attachbox: {
				show: true
			},
			attacher: {
				file: {
					popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
				}
			}
		},
		size: {
			contentWidth: 700 /* 지정된 본문영역의 넓이가 있을 경우에 설정 */
		}
	};
</script>
<!-- 에디터3 config 끝 -->

<!-- 에디터 2,3 초기화 시작 -->
<script type="text/javascript">
	EditorJSLoader.ready(function (Editor) {
		new Editor(config2);
		Editor.getCanvas().observeJob(Trex.Ev.__IFRAME_LOAD_COMPLETE, function() {
			Editor.modify({
				content: '<%=m_main_content%>'
			});
			new Editor(config3);
			Editor.getCanvas().observeJob(Trex.Ev.__IFRAME_LOAD_COMPLETE, function(ev) {
				Editor.modify({
					content: '<%=m_cmt_desc%>'
				});
			});
		});
	});
</script>
<!----- ## 모바일 폼 끝 --------------------------------->	
	<tr bgcolor="<%=adminColor("gray")%>">
		<td align="center" colspan="2">
		<% 
			'//수정
			if evt_code <> "" then 
		%>
				<input type="button" value="수정" onclick="event_reg('edit');" class="button">
		<% 
			'//신규
			else 
		%>
				<input type="button" value="등록" onclick="event_reg('add');" class="button">
		<% end if %>
				<input type="button" value="취소" onclick="self.close();" class="button">
		</td>
	</tr>
</form>	
<form name="imginputfrm" method="post" action="">
	<input type="hidden" name="YearUse" value="<%= YearUse %>">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="maxFileheight" value="">	
	<input type="hidden" name="makeThumbYn" value="">
</form>	
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->

<%
	set oip = nothing
%>