//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//	색상 선택
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	function ShowColorBox(posX, posY, gbn){
		if (ColorBox.style.visibility != 'visible'){
			ColorBox.innerHTML = tagColor(gbn);
			ColorBox.style.visibility = 'visible';
			ColorBox.style.left = (posX - 12);
			ColorBox.style.top = (posY + 12);
		}else{
			divLayerOFF();
		}
	}

	function tagColor(gb){
		//화면에 보이는 색상톤.. 
		var colortone = new Array(15);
			colortone[0] = new Array('#000000','#FFFFFF','#008000','#800000','#AC8295','#808000','#000080','#800080','#808080','#C0C0C0');
			colortone[1] = new Array('#FFFF00','#00FF00','#00FFFF','#FF00FF','#FF0000','#0000FF','#008080','#ED8602','#0099FF','#9900FF');
			colortone[2] = new Array('#FFFFFF','#E5E4E4','#D9D8D8','#C0BDBD','#A7A4A4','#8E8A8B','#827E7F','#767173','#5C585A','#000000');
			colortone[3] = new Array('#FEFCDF','#FEF4C4','#FEED9B','#FEE573','#FFED43','#F6CC0B','#E0B800','#C9A601','#AD8E00','#8C7301');
			colortone[4] = new Array('#FFDED3','#FFC4B0','#FF9D7D','#FF7A4E','#FF6600','#E95D00','#D15502','#BA4B01','#A44201','#8D3901');
			colortone[5] = new Array('#FFD2D0','#FFBAB7','#FE9A95','#FF7A73','#FF483F','#FE2419','#F10B00','#D40A00','#940000','#6D201B');
			colortone[6] = new Array('#FFDAED','#FFB7DC','#FFA1D1','#FF84C3','#FF57AC','#FD1289','#EC0078','#D6006D','#BB005F','#9B014F');
			colortone[7] = new Array('#FCD6FE','#FBBCFF','#F9A1FE','#F784FE','#F564FE','#F546FF','#F328FF','#D801E5','#C001CB','#8F0197');
			colortone[8] = new Array('#E2F0FE','#C7E2FE','#ADD5FE','#92C7FE','#6EB5FF','#48A2FF','#2690FE','#0162F4','#013ADD','#0021B0');
			colortone[9] = new Array('#D3FDFF','#ACFAFD','#7CFAFF','#4AF7FE','#1DE6FE','#01DEFF','#00CDEC','#01B6DE','#00A0C2','#0084A0');
			colortone[10] = new Array('#EDFFCF','#DFFEAA','#D1FD88','#BEFA5A','#A8F32A','#8FD80A','#79C101','#3FA701','#307F00','#156200');
			colortone[11] = new Array('#D4C89F','#DAAD88','#C49578','#C2877E','#AC8295','#C0A5C4','#969AC2','#92B7D7','#80ADAF','#9CA53B');

		var strHTML = "";
			strHTML = strHTML + "<table cellpadding=2 cellspacing=0 border=1 style='border-collapse: collapse' bgcolor='#FFFFFF'><tr><td><table cellpadding=0 cellspacing=0 border=0>";

			for (var i=0; i<11; i++){
				strHTML = strHTML + "<tr>";
				
				for(var j=0; j<10; j++){
					strHTML = strHTML + "<td onmouseover=this.style.backgroundColor='blue' onmouseout=this.style.backgroundColor='' class='hand' title='" + colortone[i][j] + "'><table cellpadding=0 cellspacing=1 border=0><tr><td bgcolor='" + colortone[i][j] + "' onclick='changeColor(\"" + colortone[i][j] + "\", "+ gb +");' width=10 height=10></td></tr></table></td>";
				}   
				strHTML = strHTML + "</tr>";
			}
			strHTML = strHTML + "</table></td></tr></table>";
			
			return strHTML;
	}
	
	function divLayerOFF(){
	
		try{
			ColorBox.style.visibility = 'hidden';
			imgLinkBox.style.visibility = 'hidden';
			LinkBox.style.visibility = 'hidden';
		}catch(e){}
	}

	function changeColor(opt, gg){
		ColorBox.style.visibility = 'hidden';
		if(gg == 1){
			frmPjt.prvColor.style.backgroundColor = opt;
			frmPjt.font1Color.value=opt;
		}else if (gg == 2){
			frmPjt.prvColor2.style.backgroundColor = opt;
			frmPjt.font2Color.value=opt;			
		}else{
			frmPjt.prvColor3.style.backgroundColor = opt;
			frmPjt.bgColor.value=opt;						
		}
	}
	document.write("<div id='ColorBox' style='position:absolute;visibility:hidden;left:200;top:100;'></div>");
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++