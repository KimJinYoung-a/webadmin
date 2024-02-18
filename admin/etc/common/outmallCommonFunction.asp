<%
Public Sub drawSelectBoxEtcLinkGbn(selectBoxName,selectedId,isDpAll)
	Dim tmp_str,query1
%>
	<select class="select" name="<%=selectBoxName%>" onchange="lgbn(this.value);">
	<% If (isDpAll) Then %>
		<option value='' <% If selectedId="" Then response.write " selected"%> >ALL</option>
	<% End If %>
<%
	query1 = " select linkgbn,valtype,linkDesc from db_item.dbo.tbl_OutMall_etcLinkGubun where 1=1 " & VBCRLF
	If poomok <> "21" Then
		query1 = query1 & " AND linkgbn <> 'infoDiv21Lotte' " & VBCRLF
	End If
	rsget.Open query1,dbget,1
	If  not rsget.EOF  Then
		Do until rsget.EOF
			If Lcase(selectedId) = Lcase(rsget("linkgbn")) Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("linkgbn")&"' "&tmp_str&">" + rsget("linkDesc") + "</option>")
			tmp_str = ""
		rsget.MoveNext
		loop
	End If
	rsget.close
	response.write("</select>")
End Sub

Public Sub drawSelectBoxXSiteAPIPartner(selBoxName, selVal)
	Dim retStr
	retStr = "<select name='"&selBoxName&"' class='select'>"
	retStr = retStr & " <option value=''>ÀüÃ¼"
	retStr = retStr & " <option value='interpark' "& CHKIIF(selVal="interpark","selected","") &" >ÀÎÅÍÆÄÅ©"
	retStr = retStr & " <option value='lotteCom' "& CHKIIF(selVal="lotteCom","selected","") &" >·Ôµ¥´åÄÄ"
	retStr = retStr & " <option value='lotteimall' "& CHKIIF(selVal="lotteimall","selected","") &" >·Ôµ¥iMall"
	retStr = retStr & " <option value='GSShop' "& CHKIIF(selVal="GSShop","selected","") &" >GSShop"
	retStr = retStr & " <option value='homeplus' "& CHKIIF(selVal="homeplus","selected","") &" >Homeplus"
	retStr = retStr & " <option value='auction1010' "& CHKIIF(selVal="auction1010","selected","") &" >¿Á¼Ç"
	retStr = retStr & " <option value='nvstorefarm' "& CHKIIF(selVal="nvstorefarm","selected","") &" >½ºÅä¾îÆÊ"
	retStr = retStr & " <option value='gmarket1010' "& CHKIIF(selVal="gmarket1010","selected","") &" >Gmarket"
	retStr = retStr & " <option value='ezwel' "& CHKIIF(selVal="ezwel","selected","") &" >ÀÌÁöÀ£"
	retStr = retStr & " <option value='coupang' "& CHKIIF(selVal="coupang","selected","") &" >ÄíÆÎ"
	retStr = retStr & " </select> "
	response.write retStr
End Sub
%>