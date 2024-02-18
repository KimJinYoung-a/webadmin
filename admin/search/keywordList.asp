<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<!-- #include virtual="/lib/classes/search/searchCls.asp" -->
<%
Dim dispCate, makerid, maxDepth, sellyn, searchstring, page, i
Dim rowNum, sortmethod, p, listKeywords, cateboostkeys, addautokeys, nvparseKeyword, sellCnt, attrNmArr, attriblist, searchkeywordlist
Dim SynonymAssign, research, idxrect, expandsearch
Dim ret_Extract, exlevel

dispCate		= requestCheckvar(request("disp"),16)
makerid			= requestCheckvar(request("makerid"),32)
sellyn			= requestCheckvar(request("sellyn"),1)
searchstring	= requestCheckvar(request("searchstring"),100)
page			= requestCheckvar(request("page"),10)
sortmethod		= requestCheckvar(request("sortmethod"),10)
SynonymAssign   = requestCheckvar(request("SynonymAssign"),10)
research        = requestCheckvar(request("research"),10)
idxrect			= requestCheckvar(request("idxrect"),16)
expandsearch	= requestCheckvar(request("expandsearch"),30)
exlevel			= requestCheckvar(request("exlevel"),10)
maxDepth	= 3

''searchstring=RepWord(searchstring,"[^°¡-ÆRa-zA-Z0-9.&%\-\_\s]","")
searchstring = RepWord(searchstring,"[^°¡-ÆRa-zA-Z0-9.&%\-\_\(\)\/\\\[\]\s]","")  
''response.write searchstring
If page = "" Then page = 1
If sortmethod = "" Then sortmethod = "bs7"
if research="" and SynonymAssign="" then SynonymAssign="on"
if research="" and expandsearch="" then expandsearch="allwordOradjacent" ''"alladjacent"
if (idxrect="") then idxrect="idx_itemname"
if (exlevel="") then exlevel="2"

Dim oDoc, itemArr, lp, oItemWord, arrList, ret_synonym
itemArr = ""
Set oDoc = new SearchItemCls
	oDoc.FCurrPage = page
	oDoc.FPageSize = 100
	oDoc.FScrollCount = 10
	oDoc.FRectSearchTxt		= searchstring
	oDoc.FRectCateCode		= dispCate				'Ä«Å×°í¸®ÄÚµå
	oDoc.FRectMakerid		= makerid				'¾÷Ã¼ ¾ÆÀÌµð
	oDoc.FListDiv			= "fulllist"
	oDoc.FSellScope			= sellyn
	oDoc.FRectSortMethod	= sortmethod
	oDoc.FRectSynonymAssign = SynonymAssign
	oDoc.FRectExpandSearch  = expandsearch
	oDoc.FRectIdxrect		= idxrect

	if ((sortmethod="bs6") or (sortmethod="be6") or (sortmethod="bs7") or (sortmethod="be7")) and (searchstring="") then
		
	else
		if ((sortmethod="bs7") or (sortmethod="be7")) and (expandsearch<>"allwordOradjacent") then

		else
			oDoc.getSearchList
		end if
	end if
	
	ret_synonym = oDoc.GetSynonymList

	ret_Extract = oDoc.ExtractKeyword(exlevel)
	
For lp = 0 To oDoc.FResultCount - 1
	itemArr = itemArr & oDoc.FItemList(lp).FItemID & ","
Next
If Right(itemArr,1) = "," Then
	itemArr = Left(itemArr, Len(itemArr) - 1)
End If

If oDoc.FResultCount = 0 Then
	'Call Alert_move("°Ë»öµ¥ÀÌÅÍ°¡ ¾ø½À´Ï´Ù.\n¸®½ºÆ®·Î ÀÌµ¿ÇÕ´Ï´Ù","/admin/search/keywordlist.asp?menupos="&menupos&"")
	'response.end
End If

Set oItemWord = new cItemContent
	arrList = oItemWord.fnItemcontents(itemArr, oDoc.FPageSize)
Set oItemWord = nothing
%>
<style>
input:-ms-input-placeholder { color: #ADADAD; }
input::-webkit-input-placeholder { color: #ADADAD; }
input::-moz-placeholder { color: #ADADAD; }
input::-moz-placeholder { color: #ADADAD; }
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}

function AnCheckClick2(e){
	if ($("#"+e+"").is(":checked")){
	}else{
		$("input:checkbox[id='"+e+"']").prop("checked", true); 
		$("#tr"+e+"").attr('class','H');
	}
}

function checkConfirmProcess() {
	var chkSel=0;
	var keywords = "";
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) {
					chkSel++;
					keywords = keywords + frmSvArr.keywords[i].value + "*(^!";
				}
			}
		} else {
			if(frmSvArr.cksel.checked){
				 chkSel++;
				 keywords = frmSvArr.keywords.value;
			}
		}
		if(chkSel<=0) {
			alert("¼±ÅÃÇÑ »óÇ°ÀÌ ¾ø½À´Ï´Ù.");
			return;
		}
	}
	catch(e) {
		alert("»óÇ°ÀÌ ¾ø½À´Ï´Ù.");
		return;
	}
	if (confirm(chkSel + '°³ÀÇ »óÇ° Å°¿öµå º¯°æÀ» Àû¿ëÇÏ½Ã°Ú½À´Ï±î?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "chk";
		document.frmSvArr.arrkeywords.value = keywords;
		document.frmSvArr.action = "/admin/search/keywordProc.asp"
		document.frmSvArr.submit();
    }
}

function popSugiConfirm() {
	var chkSel=0;
	var arritemid = "";
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) {
					chkSel++;
					arritemid = arritemid + frmSvArr.cksel[i].value + ",";
				}
			}
		} else {
			if(frmSvArr.cksel.checked){
				 chkSel++;
				 arritemid = frmSvArr.cksel.value + ",";
			}
		}
		if(chkSel<=0) {
			alert("¼±ÅÃÇÑ »óÇ°ÀÌ ¾ø½À´Ï´Ù.");
			return;
		}
	}
	catch(e) {
		alert("»óÇ°ÀÌ ¾ø½À´Ï´Ù.");
		return;
	}
    var popkeyword = window.open("/admin/search/popUpdatekeyword.asp?arritemid="+arritemid,"popkeyword","width=600,height=250,scrollbars=yes,resizable=yes");
	popkeyword.focus();
}

function pop_keywordLog(){
    var popwin = window.open("/admin/search/popkeywordLog.asp","popkeywordLog","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">°Ë»ö<br>Á¶°Ç</td>
	<td align="left">
		Àü½ÃÄ«Å×°í¸® : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
		&nbsp;&nbsp;&nbsp;
		ÆÇ¸Å¿©ºÎ : 
		<select name="sellyn" class="select">
			<option value="">ÀüÃ¼</option>
			<option value="Y" <%= Chkiif(sellyn="Y", "selected", "") %> >ÆÇ¸Å</option>
			<option value="S" <%= Chkiif(sellyn="S", "selected", "") %> >ÀÏ½ÃÇ°Àý</option>
			<option value="N" <%= Chkiif(sellyn="N", "selected", "") %> >Ç°Àý</option>
		</select>
		&nbsp;&nbsp;&nbsp;
		ÀÎµ¦½º :
		<select name="idxrect" class="select">
			<option value="idx_itemname" <%= Chkiif(idxrect="idx_itemname", "selected", "") %> >idx_itemname</option>
			<!--
			<option value="idx_keylist" <%= Chkiif(idxrect="idx_keylist", "selected", "") %> >idx_keylist</option>
			-->
		</select>
		&nbsp;&nbsp;&nbsp;
		<input type="checkbox" name="SynonymAssign" <%=CHKIIF(SynonymAssign<>"","checked","") %>>µ¿ÀÇ¾î¹Ý¿µ

		&nbsp;&nbsp;&nbsp;
		ÇüÅÂ¼ÒºÐ¼®
		<select name="exlevel">
			<option value="2" <%= Chkiif(exlevel="2", "selected", "") %> >level 2</option>
			<option value="5" <%= Chkiif(exlevel="5", "selected", "") %> >level 5</option>

			<% if (FALSE) then %>
			<option value="1" <%= Chkiif(exlevel="1", "selected", "") %> >level 1</option>
			<option value="3" <%= Chkiif(exlevel="3", "selected", "") %> >level 3</option>
			<option value="4" <%= Chkiif(exlevel="4", "selected", "") %> >level 4</option>
			<option value="6" <%= Chkiif(exlevel="6", "selected", "") %> >level 6</option>
			<% end if %>
		</select>
		<p>
		È®Àå°Ë»ö :
		<select name="expandsearch" class="select">
			<option value="" <%= Chkiif(expandsearch="", "selected", "") %> >None</option>
			<option value="allwordOradjacent" <%= Chkiif(expandsearch="allwordOradjacent", "selected", "") %> >alladjacent or allword</option>
			<option value="alladjacent" <%= Chkiif(expandsearch="alladjacent", "selected", "") %> >alladjacent</option>
			<option value="allword" <%= Chkiif(expandsearch="allword", "selected", "") %> >allword</option>
		</select>
		&nbsp;&nbsp;&nbsp;
		Á¤·Ä¹æ½Ä : 
		<select name="sortmethod" class="select">
			<option value="ne" <%= Chkiif(sortmethod="ne", "selected", "") %> >½Å»óÇ°¼ø</option>

			<option value="bs7" <%= Chkiif(sortmethod="bs7", "selected", "") %> >*ÆÇ¸Å·®¼ø(CATEGORYFIELD(recomkeyword,categorynamelist,bestkeylist ,'keyword') desc,MatchField(name1,name2) sellCnt-itemscore-itemid)</option>
			<option value="be7" <%= Chkiif(sortmethod="be7", "selected", "") %> >*ÀÎ±â»óÇ°¼ø(CATEGORYFIELD(recomkeyword,categorynamelist,bestkeylist ,'keyword') desc,MatchField(name1,name2) itemscore-itemid)</option>

			<option value="bs6" <%= Chkiif(sortmethod="bs6", "selected", "") %> >ÆÇ¸Å·®¼ø(CATEGORYFIELD(recomkeyword,categorynamelist,bestkeylist ,'keyword') desc sellCnt-itemscore-itemid)</option>
			<option value="be6" <%= Chkiif(sortmethod="be6", "selected", "") %> >ÀÎ±â»óÇ°¼ø(CATEGORYFIELD(recomkeyword,categorynamelist,bestkeylist ,'keyword') desc itemscore-itemid)</option>

			<option value="bs4" <%= Chkiif(sortmethod="bs4", "selected", "") %> >ÆÇ¸Å·®¼ø(MatchField(cate,best)-sellCNT-itemscore-itemid)</option>
			<option value="be" <%= Chkiif(sortmethod="be", "selected", "") %> >ÀÎ±â»óÇ°¼ø(MatchField(cate,best)-itemscore-itemid)</option>

			<option value="bs0" <%= Chkiif(sortmethod="bs0", "selected", "") %> >ÆÇ¸Å·®¼ø(sellCNT-sellcash-itemid)</option>

			<option value="bs1" <%= Chkiif(sortmethod="bs1", "selected", "") %> >ÆÇ¸Å·®¼ø(sellCNT-MatchField(cate,best)-itemid)</option>
			<option value="bs2" <%= Chkiif(sortmethod="bs2", "selected", "") %> >ÆÇ¸Å·®¼ø(sellCNT-MatchField(cate)-itemid)</option>
			
			<option value="bs3" <%= Chkiif(sortmethod="bs3", "selected", "") %> >ÆÇ¸Å·®¼ø(sellCNT-MatchField(cate,best)-sellcash-itemid)</option>
			
			<option value="bs5" <%= Chkiif(sortmethod="bs5", "selected", "") %> >ÆÇ¸Å·®¼ø(MatchField(cate)-sellCNT-sellcash-itemid)</option>
		</select>&nbsp;
		
		
		<br /><br />
		ºê·£µåID : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		<br /><br />
		°Ë»ö¾î :
		<input type="text" name="searchstring" size="60" class="text" value="<%=searchstring%>" placeholder="»óÇ°¸í, »óÇ°ÄÚµå, °Ë»ö Å°¿öµå¸¦ ÀÔ·ÂÇØÁÖ¼¼¿ä." onKeyPress="if (event.keyCode == 13) document.frm.submit();">
	</td>
	<td width="130">ÇüÅÂ¼ÒºÐ¼®<br><textarea cols="15" rows="3" readonly><%=ret_Extract%></textarea></td>
	<td width="350">ÇöÀçµî·ÏµÈµ¿ÀÇ¾î<br><textarea cols="45" rows="3" readonly><%=ret_synonym%></textarea></td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="°Ë»ö" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr>
	<td colspan="3" bgcolor="#FFFFFF" >
	<%=oDoc.getRetSearchQuery &"<br>"& oDoc.getRetSortQuery %>
	</td>
</tr>
</form>

</table>
<p />
<input type="button" class="button_s" value="º¯°æÀÌ·Â" onClick="javascript:pop_keywordLog();">
<p />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="arrkeywords" value="">
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="16">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				°Ë»ö°á°ú : <b><%= FormatNumber(oDoc.FTotalCount,0) %></b>
				&nbsp;
				ÆäÀÌÁö : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oDoc.FTotalPage,0) %></b>
			</td>
			<td align="right">
				<input type="button" class="button" value="ÀÏ°ýº¯°æ" onclick="popSugiConfirm();">
				&nbsp;&nbsp;<input type="button" class="button" value="Àû¿ë" onclick="checkConfirmProcess();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="40">No.</td>
	<td width="100">Àü½ÃÄ«Å×°í¸®1</td>
	<td width="100">Àü½ÃÄ«Å×°í¸®</td>
	<td width="60">»óÇ°ÄÚµå</td>
	<td width="80">ÀÌ¹ÌÁö</td>
	<td width="100">ºê·£µåID</td>
	<td>»óÇ°¸í</td>
	<td width="300">Å°¿öµå</td>
	<td width="80">BoostKey</td>
	<td width="80">EPÄ«Å×Å°¿öµå</td><!-- EP ¸ÅÇÎ»óÇ°ÀÇ Ä«Å×°í¸®¸íÀ» ÀÚµ¿À¸·Î ³ÖÀÚ -->
	<td width="80">¼Ó¼º||¿É¼Ç¼Ó¼º</td>
	<td width="50">½ºÄÚ¾î</td>
	<td width="50">ÆÇ¸Å·®</td>
	<td width="50">ÆÇ¸Å¿©ºÎ</td>
	<td width="50">»óÇ°</td>
</tr>
<%
rowNum = rowNum + (page -1) * oDoc.FPageSize +1
For i = 0 To oDoc.FResultCount - 1
	For p = 0 to ubound(arrlist,2)
		If oDoc.FItemList(i).FItemID = Trim(arrList(0,p)) Then
			listKeywords = Trim(arrList(1,p))
			cateboostkeys = Trim(arrList(2,p))
			addautokeys = Trim(arrList(3,p))
			nvparseKeyword = Trim(arrList(4,p))
			sellCnt		 = Trim(arrList(5,p))
			attrNmArr	 = Trim(arrList(6,p))
			attriblist	 = Trim(arrList(7,p))
			searchkeywordlist = Trim(arrList(8,p))
		End If
	Next
%>
<tr align="center" bgcolor="#FFFFFF" id="tr<%= oDoc.FItemList(i).FItemID %>">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" id="<%= oDoc.FItemList(i).FItemID %>"  value="<%= oDoc.FItemList(i).FItemID %>"></td>
	<td><%= rowNum %></td>
	<td>
		<%
			If Ubound(Split(oDoc.FItemList(i).FAllCateName, "^^")) > 0 Then
				rw Split(oDoc.FItemList(i).FAllCateName, "^^")(0)
			End If
		%>
	</td>
	<td>
		<%
			If Ubound(Split(oDoc.FItemList(i).FAllCateName, "^^")) > 0 Then
				rw Split(oDoc.FItemList(i).FAllCateName, "^^")(Ubound(Split(oDoc.FItemList(i).FAllCateName, "^^")))
			End If
		%>
	</td>
	<td><%= oDoc.FItemList(i).FItemID %></td>
	<td><img src="<%= oDoc.FItemList(i).FImageSmall %>"></td>
	<td><%= oDoc.FItemList(i).FMakerid %></td>
	<td><%= oDoc.FItemList(i).FItemname %></td>
	<td>
	    <input type="text" name="keywords" class="text" size="70" onclick="AnCheckClick2('<%=oDoc.FItemList(i).FItemID%>');" value="<%= listKeywords %>">
	    <% if (searchkeywordlist<>"") then %><br><%=replace(searchkeywordlist,searchstring,"<strong>"&searchstring&"</strong>")%><% end if %>		
	</td>
	<td><%=replace(cateboostkeys,searchstring,"<strong>"&searchstring&"</strong>")%></td>
	<td><%=replace(nvparseKeyword,searchstring,"<strong>"&searchstring&"</strong>")%></td>
	<td><%=attrNmArr%>||<%=attriblist%></td>
	<td><%= oDoc.FItemList(i).FItemscore %></td>
	<td><%= FormatNumber(sellCnt,0) %></td>
	<td>
	    <% if oDoc.FItemList(i).FSellyn<>"Y" then %>
	    <strong><%= oDoc.FItemList(i).FSellyn %></strong>
	    <% else %>
	    <%= oDoc.FItemList(i).FSellyn %>
	    <% end if %>
	</td>
	<td><a href="<%=wwwURL%>/<%= oDoc.FItemList(i).FItemID %>" target="_blank">GO ></a></td>
</tr>
<%
	rowNum = rowNum + 1 
Next
%>
<% if oDoc.FResultCount<1 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center"> 
	<% if ((sortmethod="bs6") or (sortmethod="be6") or (sortmethod="bs7") or (sortmethod="be7")) and (searchstring="") then %>
	Á¤·Ä¹æ½ÄÀÌ <strong>CATEGORYFIELD</strong> ÀÎ °æ¿ì °Ë»ö¾î°¡ ÇÊ¿äÇÕ´Ï´Ù.
	<% elseif ((sortmethod="bs7") or (sortmethod="be7")) and (expandsearch<>"allwordOradjacent") then %>
	aliasing À» »ç¿ëÇÏ´Â°æ¿ì È®Àå°Ë»ö¿¡ aliasing À» ¼±¾ðÇØ¾ßÇÕ´Ï´Ù.
	<% else %>
	°Ë»ö°á°ú°¡ ¾ø½À´Ï´Ù.
	<% end if %>
	</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">
	<% If oDoc.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oDoc.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oDoc.StartScrollPage To oDoc.FScrollCount + oDoc.StartScrollPage - 1 %>
		<% If i>oDoc.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oDoc.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set oDoc = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->