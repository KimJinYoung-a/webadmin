<%@ language="VBScript" %>
<% option explicit %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
  
<%
	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode
	Dim nIndent, strtitle
	Dim nInstr,searchmode,search,searchString
	dim iboard_idx
    Dim nboard,arrFile,intF
	dim sRegType : sRegType = "A"
	if Request("pgsize")="" then
		pgsize = 10
	else
		pgsize = Request("pgsize")
	end if
	
	if Request("page") = "" then
		page = 1
	else
		page = cInt(Request("page")) 
	end if

iboard_idx = requestCheckvar(request("idx"),10)
set nboard = new CBoard
nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
nboard.design_notice_read iboard_idx
nboard.FRectidx		 = iboard_idx
	arrFile   = nboard.fnGetAttachFile 
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script> 
<script language="JavaScript">
<!--
function gotolist(){
location.href = "designer_notice.asp?idx=<%= request("idx") %>&page=<% =page %>&menupos=79"
}
function gotomodify(){
location.href = "designer_notice_modify.asp?idx=<%= request("idx") %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=79"
}

//파일 다운로드
    function jsDownload(sDownURL, sRFN, sFN){
    var winFD = window.open(sDownURL+"/linkweb/board/procDownload.asp?sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
 }
//-->
</script>

<table width="800"   cellpadding="3" cellspacing="1" class="a" >
	<tr>
		<td>
			<table width="100%"   cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	    <tr> 
	      <td height="25" valign="middle" bgcolor="<%= adminColor("tabletop") %>">  
	      	 <b>  <%=nboard.FRectTitle %></b> 
	      </td>
	    </tr>
	    <tr bgcolor="#FFFFFF"> 
	    	<td>
	    		<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a">
		    		<tr> 
		      			<td> 아이디: <span class="id"><% =nboard.FRectID %></span> &nbsp;|
		    			  글쓴이: <span class="id"><%=nboard.FRectName %></span>&nbsp;| 날짜: <% =(nboard.Fregdate) %> | <%if nboard.Fdispcatename <> "" then%>카테고리: <%=nboard.Fdispcatename%><%end if%></td>
				    </tr>
				    <tr> 
				      <td><img src="/admin/images/w_dot.gif" width="100%" height="1"></td>
				    </tr>
				     <tr> 
				      <td valign="top" bgcolor="#FFFFFF" height="500"> 
				        내용 :<br>
				         <%=nboard.FRectContents %>
				          <br>
				      </td>
				    </tr>
				    <tr> 
				    <td height="2"><img src="/admin/images/w_dot.gif" width="100%" height="1"></td>
				    </tr>
				    <tr> 
				    	<td>첨부파일:
				    		<div id="dFile">
						<% Dim arrFName,arrF, sFName, intF2,intF3, iCount 
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2) 
					
								arrF = split(arrFile(2,intF),"/") 
							 	arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0)  
						%>
						<div id="dF<%=sFName%>"><a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a>&nbsp;<input type="button" value="x" class="button" onclick="jsFileDel('<%=sFName%>')"> 
						 </div>
					<%Next
						END IF
						%> 
						</div> 
				    	</td>
				    </tr>
				    
					</table>
				</td>
			</tr>
 		</table>
 </td>
</tr>
<tr>
	<td align="center" >
		<input type="button" class="button"  value="리스트" onclick="gotolist();">&nbsp;
		<input type="button" class="button" value="글 수정" onclick="gotomodify();">
	</td>
</tr>
<tr>
	<td>
		<div style="padding:20 0 10 0"> 
	 		<!-- #include virtual="/admin/board/incComment.asp"--> 
		</div>	
	</td>
</tr>
</table>
 
<%
set nboard = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
 
