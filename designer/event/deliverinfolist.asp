<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
'' 택배사 일괄적용
Sub drawSelectBoxDeliverCompanyAssign(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select name="<%=selectBoxName%>" onChange="AssignDeliverSelect(this);">
     <option value='' <%if selectedId="" then response.write " selected"%>>택배사선택</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("divcd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub


dim makerid, page, notfinish, research
makerid     = session("ssBctID")
page        = requestCheckVar(request("page"),10)
notfinish   = requestCheckVar(request("notfinish"),10)
research    = requestCheckVar(request("research"),10)

if page="" then page=1
if research="" and notfinish="" then notfinish="on"

dim oevtsongjang
set oevtsongjang = new CEventsBeasong
oevtsongjang.FPageSize           = 100
oevtsongjang.FCurrPage           = page
oevtsongjang.FRectOnlySongjangNotInput = notfinish
oevtsongjang.FRectDeleteyn       = "N"
'oevtsongjang.FRectDeliverAreaInputedOnly    = "Y"
oevtsongjang.FRectIsupchebeasong = "Y"
oevtsongjang.FRectDeliverMakerid = makerid

if (makerid<>"") then
    oevtsongjang.getEventBeasongInfoList
end if

dim i

%>
<script language='javascript'>
function AssignDeliverSelect(comp){
    var frm = comp.form;
	var selecidx = comp.selectedIndex;
	var frm;
    
    if (frm.chkidx.length>1){
    	for (var i=0;i<frm.songjangdiv.length;i++){
    	    frm.songjangdiv[i][selecidx].selected=true;
    	}
    }else{
        frm.songjangdiv[selecidx].selected=true;
    }
}

function switchCheckBox(comp){
    var frm = comp.form;
    
	if(frm.chkidx.length>1){
		for(i=0;i<frm.chkidx.length;i++){
		    if (!frm.chkidx[i].disabled){
    			frm.chkidx[i].checked = comp.checked;
    			AnCheckClick(frm.chkidx[i]);
    		}
		}
	}else{
	    if (!frm.chkidx.disabled){
    		frm.chkidx.checked = comp.checked;
    		AnCheckClick(frm.chkidx);
    	}
	}
}

function CheckThis(comp,i){
    var frm = comp.form;
    
	if (comp.value.length>5){
	    if (frm.songjangno.length>1){
	        frm.chkidx[i].checked=true;
	        AnCheckClick(frm.chkidx[i]);
        }else{
            frm.chkidx.checked=true;
            AnCheckClick(frm.chkidx);
        }
	}
}

function ReSearch(frm){
    frm.submit();
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function ShowDeliverInfo(iid){
    var popwin = window.open('popeventdeliverInfo.asp?id=' + iid,'ShowDeliverInfo','width=600,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function printBalju(){
    var frm = document.frmbalju;
    var isChecked = false;
    
    if(frm.chkidx.length>1){
		for(i=0;i<frm.chkidx.length;i++){
			if (frm.chkidx[i].checked){
			    isChecked = true;
			    break;
			}
		}
	}else{
		isChecked = frm.chkidx.checked ;
	}
	
	if (!isChecked){
	    alert('선택 내역이 없습니다.');
	    return;
	}
	
	var popwin = window.open('_blank','printBalju','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
	
	
	frm.target = 'printBalju';
    frm.action="popEventBeasongBaljulist.asp"
	frm.submit();
	
}

function saveSongjang(){
    var frm = document.frmbalju;
    var isChecked = false;
    
    if(frm.chkidx.length>1){
		for(i=0;i<frm.chkidx.length;i++){
			if (frm.chkidx[i].checked){
			    if (frm.songjangdiv[i].value.length<1){
			        alert('택배사를 선택하세요.');
			        frm.songjangdiv[i].focus();
			        return;
			    }
			    
			    if (frm.songjangno[i].value.length<6){
			        alert('송장 번호를 정확히 입력 하세요.');
			        frm.songjangno[i].focus();
			        return;
			    }
			    
			    isChecked = true;
			}
		}
	}else{
	    if (frm.songjangdiv.value.length<1){
	        alert('택배사를 선택하세요.');
	        frm.songjangdiv.focus();
	        return;
	    }
	    
	    if (frm.songjangno.value.length<6){
	        alert('송장 번호를 정확히 입력 하세요.');
	        frm.songjangno.focus();
	        return;
	    }
			    
		isChecked = frm.chkidx.checked ;
	}
	
	if (!isChecked){
	    alert('선택 내역이 없습니다.');
	    return;
	}
    
    
    if (confirm('선택 내역을 완료 처리 하시겠습니까?')){
        // not include InChecked Param
        if(frm.chkidx.length>1){
    		for(i=0;i<frm.chkidx.length;i++){
    		    frm.songjangdiv[i].disabled = (!frm.chkidx[i].checked);
    			frm.songjangno[i].disabled = (!frm.chkidx[i].checked);
    		}
    	}else{
    		isChecked = frm.chkidx.checked ;
    		frm.songjangdiv.disabled = (!frm.chkidx.checked);
    		frm.songjangno.disabled = (!frm.chkidx.checked);
    	}
    	frm.target="";
	    frm.action="doeventsongjanginput.asp"
        frm.submit();
    }
}



</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="T">
	<input type="hidden" name="page" value="">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<input type="checkbox" name="notfinish" <% if notfinish="on" then response.write "checked" %> >미출고만 검색
        </td>
        <td align="right">
			<a href="javascript:ReSearch(frm)"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr>
        <td colspan="10" bgcolor="#FFFFFF">
            <input type="button" class="button" value="선택 내역 발주서 출력" onclick="printBalju();" onFocus="this.blur();">
            &nbsp;&nbsp;
            <input type="button" class="button" value="선택 내역 송장 입력" onclick="saveSongjang();" onFocus="this.blur();">
        </td>
    </tr>
    <form name="frmbalju" method="post" action="doeventsongjanginput.asp">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30"><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
		<td>구분명</td>
		<td width="70">아이디</td>
		<td width="60">고객명</td>
		<td width="60">수령인</td>
		<td>상품명</td>
		<td width="80"><!-- (배송지)-->등록일</td>
		<td width="130"><% drawSelectBoxDeliverCompanyAssign "defaultsongjangdiv","" %></td>
		<td width="110">운송장번호</td>
		<td width="90">출고요청일<br>출고일</td>
	</tr>
	<% for i=0 to oevtsongjang.FResultCount-1 %>
	
	<% if oevtsongjang.FItemList(i).Fdeleteyn="Y" then %>
	<tr bgcolor="#CCCCCC" >
	<% else %>
	<tr bgcolor="#FFFFFF" >
	<% end if %>
	    
		<td align="center">
		    <% if IsNULL(oevtsongjang.FItemList(i).Finputdate) then %>
		    <input type="checkbox" name="chkidx" onClick="AnCheckClick(this);" value="<%= oevtsongjang.FItemList(i).Fid %>" disabled >
		    <% else %>
		    <input type="checkbox" name="chkidx" onClick="AnCheckClick(this);" value="<%= oevtsongjang.FItemList(i).Fid %>">
		    <% end if %>
		</td>
		<td align="center"><a href="javascript:ShowDeliverInfo('<%= oevtsongjang.FItemList(i).Fid %>');"><%= oevtsongjang.FItemList(i).Fgubunname %></a></td>
		<td align="center"><%= printUserId(oevtsongjang.FItemList(i).FUserId,2,"**") %></td>
		<td align="center"><%= oevtsongjang.FItemList(i).Fusername %></td>
		<td align="center"><%= oevtsongjang.FItemList(i).FReqName %></td>
		<td align="center"><%= oevtsongjang.FItemList(i).getPrizeTitle %></td>
		<td align="center">
		    <% if IsNULL(oevtsongjang.FItemList(i).Finputdate) then %>
		        <font color="red">(배송지 미입력)</font>
		    <% else %>
    		    <%= FormatDateTime(oevtsongjang.FItemList(i).Finputdate,2) %>
    		<% end if %>
		</td>
		<td align="center">
		<% drawSelectBoxDeliverCompany "songjangdiv",oevtsongjang.FItemList(i).FSongjangdiv %>
		</td>
		<td align="center">
		    <input type="text" class="text" name="songjangno" size="12" value="<%= oevtsongjang.FItemList(i).FSongjangno %>" onKeyup="CheckThis(this,'<%= i %>');" maxlength=16>
		</td>
		<td align="center">
		<% if oevtsongjang.FItemList(i).FreqDeliverDate<> "" then %>
		    (<%= oevtsongjang.FItemList(i).FreqDeliverDate %>)<br>
		<% end if %>
		
		<% if (oevtsongjang.FItemList(i).Fsenddate <> "") then %>
		    <% = FormatDateTime(oevtsongjang.FItemList(i).Fsenddate,2) %>
		<% else %>
		    <font color="red">미출고</font>
		<% end if %>
		</td>
	</tr>
	
	<% next %>
	</form>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if oevtsongjang.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oevtsongjang.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + oevtsongjang.StartScrollPage to oevtsongjang.FScrollCount + oevtsongjang.StartScrollPage - 1 %>
				<% if i>oevtsongjang.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if oevtsongjang.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    <!--
    <form name="frmArrupdate" method="post" >
	<input type="hidden" name="idarr" value="">
	</form>
	-->
</table>
<%
set oevtsongjang = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->