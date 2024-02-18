<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  핑거스 강사 게시판
' History : 2010.03.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/board/lecturer/lecturer_cls.asp"-->

<%
Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Type, sDoc_Import, sDoc_part_sn
dim sDoc_Diffi, sDoc_Subj, sDoc_Content ,sDoc_UseYN, sDoc_Regdate
dim sDoc_WorkerView , i , tContents , olect, lectFile , arrFileList ,g_MenuPos
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	g_MenuPos = RequestCheckvar(request("menupos"),10)	
	
	If iDoc_Idx = "" Then
		sDoc_Id 		= session("ssBctId")
		sDoc_Name		= session("ssBctCname")
		sDoc_Regdate	= Left(now(),10)
		sDoc_Status = "K001"
	Else
		
		Set olect = New clecturer_list
		olect.FrectDoc_Idx = iDoc_Idx
		olect.fnGetlecturerView
	
		sDoc_Id 		= olect.FOneItem.FDoc_Id
		sDoc_Name		= olect.FOneItem.FDoc_Name
		sDoc_Status		= olect.FOneItem.FDoc_Status
		if sDoc_Status = "" then sDoc_Status = "K001"			
		sDoc_Type		= olect.FOneItem.FDoc_Type
		sDoc_Import		= olect.FOneItem.FDoc_Import
		sDoc_Diffi		= olect.FOneItem.FDoc_Diffi
		sDoc_Subj		= olect.FOneItem.FDoc_Subj
		tContents	= olect.FOneItem.FDoc_Content	
		sDoc_UseYN		= olect.FOneItem.FDoc_UseYN
		sDoc_Regdate	= olect.FOneItem.FDoc_Regdate
		sDoc_part_sn	= olect.FOneItem.fpart_sn

		set lectFile = new clecturer_list
	 	lectFile.FrectDoc_Idx = iDoc_Idx
		arrFileList = lectFile.fnGetFileList	
	End If

if sDoc_Type = "" then sDoc_Type = "G010"
if sDoc_Import = "" then sDoc_Import = "L001"
%>
<!-- daumeditor head ------------------------->
<% if (FALSE) then %>
 <meta http-equiv="Content-Type" content="text/html" />
 <meta http-equiv="X-UA-Compatible" content="IE=10" /> 
<% end if %>
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

<script type="text/javascript">

var blockChar=["&lt;script","<scrip","<form","&lt;form"];  
function chkContent(p) {
	for (var i=0; i<blockChar.length; i++) {
		if (p.indexOf(blockChar[i])>=0) {
			return blockChar[i];
		}
	}
	return null;
}

function fileupload(){
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

function checkform(frm){

	if (frm.G000.value == ""){
		alert("업무 구분을 선택해 주세요!");
		return;
	}

	count = 0;
	num = frm.L000.length;
	
	for(i=0; i<num; i++){
		if(frm.L000[i].checked == true)
		{
			count +=1;
		}
	}
	if(count==0){
		alert("업무 중요도를 선택해 주세요!");
		return;
	}
	
	if (frm.doc_subject.value == ""){
		alert("제목을 입력해 주세요!");
		frm.doc_subject.focus();
		return;
	}

    var content = Editor.getContent();
    	var str = chkContent(content); 
		  if (str) {
		   alert("script태그 또는 form태그는 사용할 수 없는 문자열 입니다.");
		   return ;  
		  } 
    document.getElementById("brd_content").value = content; 
	
	if(frm.brd_content.value==''){				
		alert('내용을 입력해주세요')
		frm.brd_content.focus();
		return;
	}			
	
	frm.submit();
}

</script>

<form name="frm" action="lecturer_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<input type="hidden" name="gubun" value="write">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="doc_difficult" value="2">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">

<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<tr>
	<td style="padding-bottom:10"> 
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<% If iDoc_Idx <> "" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번호</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=iDoc_Idx%></td>
		</tr>
		<input type="hidden" name="doc_useyn" value="<%=sDoc_UseYN%>">
		<% End If %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<%= getthefingers_staff("", sDoc_part_sn, sDoc_Name) %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%=sDoc_Regdate%>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">현재 상태</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("w","K000",sDoc_Status)%></td>
		</tr>		
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">구분</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td><%=CommonCode("w","G000",sDoc_Type)%></td>					
				</tr>
				</table>
				<div id="yyy0" style="background-color:white; border-width:1px; border-style:solid; width:270; height:50; position:absolute; left:10; top:10; z-index:1; display:none"></div>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">중요도</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("w","L000",sDoc_Import)%></td>
		</tr>		
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="doc_subject" value="<%=sDoc_Subj%>" size="95" maxlength="148">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내 용</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<textarea name="brd_content" id="brd_content" style="width: 100%; height: 490px;"><%=replace(tContents,"</p><p>&nbsp;</p>","<BR></p>")%></textarea>  
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
		<tr height="30">
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
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="button" onclick="checkform(frm);" value="저장하기" class="button">
		<input type="button" value="목록으로" onclick="location.href='lecturer.asp?menupos=<%=g_MenuPos%>'" class="button">		
	</td>	
</tr>
</table>

</form>

<% If iDoc_Idx <> "" Then %>
<!-- ####### 답변쓰기 ####### //-->
<br>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td>
		<img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>답변</b>
	</td>
</tr>
</table>
<iframe src="iframe_lecturer_ans.asp?didx=<%=iDoc_Idx%>" name="iframeDB1" width="814" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<!-- ####### 답변쓰기 ####### //-->
<% End If %>

<%	
set olect = nothing
set lectFile = nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
