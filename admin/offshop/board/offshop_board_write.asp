<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  오프라인 통합 게시판
' History : 2010.06.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/board/board_cls_utf8.asp"-->

<%
Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Type, sDoc_Import , shopidcount ,dispshopdivon ,dispshopidon
dim sDoc_Diffi, sDoc_Subj, sDoc_Content ,sDoc_UseYN, sDoc_Regdate ,oread ,dispshopall ,dispshopdiv ,oshop ,doc_kind
dim i , tContents , olect, lectFile , arrFileList ,g_MenuPos ,sDoc_WorkerName ,sDoc_ViewList,sDoc_WorkerView
	iDoc_Idx = requestCheckVar(Request("didx"),10)
	g_MenuPos = requestCheckVar(request("menupos"),10)

If iDoc_Idx = "" Then
	sDoc_Id 		= session("ssBctId")
	sDoc_Name		= session("ssBctCname")
	sDoc_Regdate	= Left(now(),10)
	sDoc_Status = "01"	
Else
	
	'//상세정보
	Set olect = New clecturer_list
		olect.FrectDoc_Idx = iDoc_Idx
		olect.fnGetlecturerView()

		sDoc_Id 		= olect.FOneItem.FDoc_Id
		sDoc_Name		= olect.FOneItem.Fusername
		sDoc_Status		= olect.FOneItem.FDoc_Status
		if sDoc_Status = "" then sDoc_Status = "01"			
		sDoc_Type		= olect.FOneItem.FDoc_Type
		sDoc_Import		= olect.FOneItem.FDoc_Import
		sDoc_Diffi		= olect.FOneItem.FDoc_Diffi
		sDoc_Subj		= ReplaceBracket(olect.FOneItem.FDoc_Subj)
		tContents	= ReplaceBracket(replace(replace(olect.FOneItem.FDoc_Content,vbcrlf,"<br>"),"'",""))
		sDoc_UseYN		= olect.FOneItem.FDoc_UseYN
		sDoc_Regdate	= olect.FOneItem.FDoc_Regdate
		shopidcount = olect.FOneItem.fshopidcount
		dispshopall = olect.FOneItem.fdispshopall
		dispshopdiv = olect.FOneItem.fdispshopdiv
		doc_kind = olect.FOneItem.fdoc_kind
		
		if shopidcount > 0 then dispshopidon = "ON"
		if dispshopdiv <> "" and not isnull(dispshopdiv) then dispshopdivon = "ON"
		
	'/글 확일 날짜 저장		'/본인이면 제낌
	if session("ssBctId") <> sDoc_Id then
		Call WorkerView(iDoc_Idx)
	end if
	
	'//글 확인한 날짜 리스트
	Set oread = New clecturer_list
		oread.FrectDoc_Idx = iDoc_Idx
		oread.fnGetlecturerread()

		sDoc_WorkerName	= oread.FDoc_WorkerName
		sDoc_WorkerView	= oread.FDoc_WorkerViewdate	
	
	'/첨부파일 리스트
	set lectFile = new clecturer_list
	 	lectFile.FrectDoc_Idx = iDoc_Idx
		arrFileList = lectFile.fnGetFileList	
	
	'//위탁매장 리스트
    set oshop = new clecturer_list
    oshop.FrectDoc_Idx = iDoc_Idx
    
    '/위탁매장이 있을경우에만 쿼리
    if shopidcount > 0 then
    	oshop.getShopList
    end if
		    
	'/확인일이 있는경우에만
	For i=0 To UBOUND(Split(sDoc_WorkerName,","))
		if Not(sDoc_WorkerView="" or isNull(sDoc_WorkerView)) then
			sDoc_ViewList = sDoc_ViewList & "&nbsp;" & Split(sDoc_WorkerName,",")(i) & " : " & Split(sDoc_WorkerView,",")(i) & "<br>"
		end if
	Next		
End If
	
if sDoc_Import = "" then sDoc_Import = "02"
%>

<style>
	.display_date { cursor:pointer; display:inline-block; width:80px; border:1px solid; border-color:#a6a6a6 #d8d8d8 #d8d8d8 #a6a6a6; height:1em; padding:1px; }
</style>
<!-- daumeditor head ------------------------->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=10" /> 
<link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="utf-8"/>
<script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="utf-8"></script>
<script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
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

<script type="text/javascript">

function fileupload()
{
	window.open('popUpload.asp','worker','width=420,height=200,scrollbars=yes');
}

function clearRow(tdObj) {
	if(confirm("선택하신 파일을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;
	
		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}

//매장구분 선택시
function chdispshopdiv(){
	
	//alert(frm.dispshopdivon.checked);
	if (frm.dispshopdivon.checked){
		divdispshopdiv.style.display = '';
	}else{
		divdispshopdiv.style.display = 'none';
	}
}

//위탁매장 선택시
function chdispshopiddiv(){
	
	//alert(frm.dispshopdivon.checked);
	if (frm.dispshopidon.checked){
		dispshopiddiv.style.display = '';
	}else{
		dispshopiddiv.style.display = 'none';
	}
}

//구분지정에 따른 매장지정 변경
function chdoc_status(tmp){
	
	//업무협조 일때
	if (tmp == '02'){
		divdoc_status.style.display = '';
		divdispshop.style.display = 'none';		
		divdispshop90.style.display = '';
	} else {
		divdoc_status.style.display = 'none';
		divdispshop.style.display = '';		
		divdispshop90.style.display = 'none';
	}
}

//팝업에서 매니져 선택 추가
function addSelectedShop(shopid,shopname){
	var lenRow = tableshop.rows.length;

	// 기존에 값에 중복값 여부 검사
	if(lenRow>1)	{
		for(l=0;l<document.all.shopid.length;l++)	{
			if(document.all.shopid[l].value==shopid) {
				alert("이미 지정된 매장 입니다");
				return;
			}
		}
	}
	else {
		if(lenRow>0) {
			if(document.all.shopid.value==shopid) {
				alert("이미 지정된 매장 입니다");
				return;
			}
		}
	}

	// 행추가
	var oRow = tableshop.insertRow(lenRow);
	oRow.onmouseover=function(){tableshop.clickedRowIndex=this.rowIndex};

	// 셀추가 (이름,삭제버튼)
	var oCell1 = oRow.insertCell(0);		
	var oCell3 = oRow.insertCell(1);

	oCell1.innerHTML = shopname + "<input type='hidden' name='shopid' value='" + shopid + "'>";
	oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdshopid()' align=absmiddle>";
}

// 선택삭제
function delSelectdshopid(){
    
	if(confirm("선택한 매장을 삭제하시겠습니까?"))
		tableshop.deleteRow(tableshop.clickedRowIndex);
}

// 매장 선택 팝업
function popShopSelect(){
	var popwin = window.open("/common/offshop/pop_shopselect_pos.asp", "popShopSelect","width=600,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

//전송
function checkform(frm){
	var count;
	var num;
	
	if (frm.G000.value == "")
	{
		alert("구분을 선택해 주세요!");
		frm.G000.focus();
		return;
	}	
	
	if (frm.G000.value != '02'){
		if (frm.dispshopall.checked == false && frm.dispshopdivon.checked == false && frm.dispshopidon.checked == false){
			alert("매장지정을 선택해주세요!");
			return;
		}
	}
	
	count = 0;
	num = frm.A000.length;
	if (frm.dispshopdivon.checked == true){
		for(i=0; i<num; i++)
		{
			if(frm.A000[i].checked == true)
			{
				count +=1;
			}
		}
		if(count==0)
		{
			alert("매장구분을 선택해주세요!");
			return;
		}
	}

	if (frm.doc_kind.value == "")
	{
		alert("종류를 선택해 주세요!");
		return;
	}
	
	count = 0;
	num = frm.L000.length;
	for(i=0; i<num; i++)
	{
		if(frm.L000[i].checked == true)
		{
			count +=1;
		}
	}
	if(count==0)
	{
		alert("업무 중요도를 선택해 주세요!");
		return;
	}
	
	if (frm.doc_subject.value == "")
	{
		alert("제목을 입력해 주세요!");
		frm.doc_subject.focus();
		return;
	}

    var content = Editor.getContent();
	var str = chkContent(content); 
	if (str) {
		alert("script태그, iframe태그 , form태그는 사용할 수 없는 문자열 입니다.");
		return ;  
	} 
    document.getElementById("brd_content").value = content; 

	if(frm.doc_content.value==''){				
		alert('내용을 입력해주세요')
		return;
	}			
	
	frm.submit();
}

var blockChar=["&lt;script","<scrip","<form","&lt;form","<iframe"];  
function chkContent(p) {
	for (var i=0; i<blockChar.length; i++) {
		if (p.indexOf(blockChar[i])>=0) {
			return blockChar[i];
		}
	}
	return null;
}

</script>

※본사공지 : 본사에서 각 매장에 알리는 공지사항 입니다.(본사에서만 작성가능)
<br>업무협조 : 각매장에서 본사에 요청하는 글입니다.(매장에서만 작성가능)
<br>매장공지 : 각매장에서 타 매장에 알리는 공지사항입니다.(매장에서만 작성가능)
<form name="frm" action="/admin/offshop/board/offshop_board_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<input type="hidden" name="gubun" value="write">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="doc_difficult" value="2">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">
<table border="0" cellpadding="0" cellspacing="0" class="a" width="814">
<tr>
	<td align="right">
		<input type="button" value="목록으로" onclick="location.href='offshop_board.asp?menupos=<%=g_MenuPos%>'" class="button">		
	</td>	
</tr>
<tr>
	<td style="padding-bottom:10"> 
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" width="100%">
		<% If iDoc_Idx <> "" Then %>
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번호</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=iDoc_Idx%></td>
		</tr>
		<input type="hidden" name="doc_useyn" value="<%=sDoc_UseYN%>">
		<% End If %>
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Name%>(<%=sDoc_Id%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%=sDoc_Regdate%></td>
		</tr>
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">구분</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<%
				'//매장에서 쓴글일경우 본사에서 해당내역을 볼수는 있지만 구분을 수정 할수는 없음
				if C_ADMIN_USER and (sDoc_Type = "02" or sDoc_Type = "03") then 
				%>
					<%=CommonCode("v","G000",sDoc_Type,C_ADMIN_USER," onchange='chdoc_status(this.value);'") %>
					<input type="hidden" name="G000" value="<%=sDoc_Type%>">
				<% else %>
					<%=CommonCode("w","G000",sDoc_Type,C_ADMIN_USER," onchange='chdoc_status(this.value);'") %>
				<% end if %>
				<!-- 업무협조 일때만 뿌림 -->	
				<div id="divdoc_status" name="divdoc_status" style="display:<% if sDoc_Type <> "02" then response.write " none"%>">
					현재 상태 : <%=CommonCode("w","K000",sDoc_Status,"","")%>
				</div>
			</td>
		</tr>
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">매장지정</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table border="0" cellspacing="0" class="a">
				<!--업무협조 일경우 뿌림-->
				<tr id="divdispshop90" name="divdispshop90" style="display:<% if sDoc_Type <> "02" then response.write " none" %>">
			        <td>
						본사
			        </td>
			    </tr>
			    <!--업무협조가 아닌경우 뿌림-->
			    <tr id="divdispshop" name="divdispshop" style="display:<% if sDoc_Type = "02" then response.write " none" %> ">
			    	<td >
				    	<table border="0" cellspacing="0" class="a" width="100%">
				    	<tr>
					        <td colspan=2>
					        	<input type="checkbox" name="dispshopall" value="ON" <% if dispshopall="Y" then response.write " checked" %>>전체매장(직영+가맹+해외) / 기타매장은 위탁매장에서 지정하세요.
		        			</td>
		        		</tr>
				    	<tr>
					        <td >
								<input type="checkbox" name="dispshopdivon" value="ON" <% if dispshopdivon="ON" then response.write " checked" %> onclick="chdispshopdiv();">매장구분	        			
		        			</td>
					        <td >
								<div id="divdispshopdiv" name="divdispshopdiv" style="display:<% if dispshopdivon<>"ON" then response.write " none" %> ">
									&nbsp;&nbsp;&nbsp;&nbsp;<%=CommonCode("w","A000",dispshopdiv,"","")%>
								</div>					        			
		        			</td>		        			
		        		</tr>		        							        			    		
				    	<tr>
					        <td valign="top">
								<input type="checkbox" name="dispshopidon" value="ON" <% if dispshopidon="ON" then response.write " checked" %> onclick="chdispshopiddiv();">위탁매장
					        </td>
		        			<td valign="bottom">
								<table border="0" cellspacing="0" class="a">
								<tr id="dispshopiddiv" name="dispshopiddiv" style="display:<% if dispshopidon <> "ON" then response.write " none" %> ">
									<td>
									    <table name='tableshop' id='tableshop' class=a>
									        <%
									        if iDoc_Idx <> "" then
									        	if oshop.FResultCount > 0 then
									        
									        	for i=0 to oshop.FResultCount-1
									        %>
										        <tr onMouseOver='tableshop.clickedRowIndex=this.rowIndex'>
											    	<td>
											    	    <%= oshop.FItemList(i).fshopname %>
											    	    <input type='hidden' name='shopid' value='<%= oshop.FItemList(i).fshopid %>'></td>  
											    	<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdshopid()' align=absmiddle></td>   
										        </tr>
									        <%
									        	next
									        
									    		end if
									    	end if
									        %>
									    </table>
									</td>
									<td valign="bottom"><input type="button" class='button' value="매장추가" onClick="popShopSelect()"></td>
								<tr>
								</table>
		        			</td>
		        		</tr>
	        			</table>
        			</td>
        		</tr>
        		</table>
			</td>
		</tr>
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<%=CommonCode("w","doc_kind",doc_kind,C_ADMIN_USER,"") %>
			</td>
		</tr>		
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">중요도</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("w","L000",sDoc_Import,"","")%></td>
		</tr>			
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="doc_subject" value="<%=sDoc_Subj%>" size="95" maxlength="148">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내 용</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<br><font color="red">다른곳의 글을 HTML 태그까지 모두 복사해서 붙여 넣을경우 에러가 발생할수 있습니다. 참고 하세요</font>
				<textarea name="doc_content" id="brd_content" style="width: 100%; height: 490px;"><%=replace(tContents,"</p><p>&nbsp;</p>","<BR></p>")%></textarea>  
                <!-- daumeditor  --> 
                <script type="text/javascript">  
                    EditorCreator.convert(document.getElementById("brd_content"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                        EditorJSLoader.ready(function (Editor) {
                            new Editor(config);
                            Editor.modify({
                                content: document.getElementById("brd_content")
                            });
                        });
                    });
                
                </script> 
                <!-- daumeditor   -->
			</td>
		</tr>
		<tr >
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td width="100" valign="top" style="padding:5 0 0 0">
						<input type="button" value="파일업로드" onClick="fileupload()" class="button"> 
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
						(주민번호등 개인정보가 있는 파일은 올리시면 안됩니다.)
					</td>
				</tr>
				
				</table>
			</td>
		</tr>
		<% if sDoc_ViewList <> "" then %>
			<tr >
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">글확인<br>명단</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%= sDoc_ViewList %>
				</td>
			</tr>
		<% end if %>	
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="button" onclick="checkform(frm);" value="저장하기" class="button">		
	</td>	
</tr>
</table>
</form>

<% If iDoc_Idx <> "" Then %>
<!-- ####### 답변쓰기 ####### //-->
<br>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr >
	<td>
		<img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>답변</b>
	</td>
</tr>
</table>
<iframe src="iframe_board_ans.asp?didx=<%=iDoc_Idx%>" name="iframeDB1" width="814" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<!-- ####### 답변쓰기 ####### //-->
<% End If %>

<%
session.codePage = 949

set olect = nothing
set lectFile = nothing
set oread = nothing
set oshop = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
