 
var theobject = null;         
var theobject2 = null;  

function resizeObject() {
        this.el        = null; 
        this.dir    = "";     
        this.grabx = null;    
        this.width = null; 
        this.left = null; 
}
        

function getDirection(el) {
        var xPos,  offset, dir;
        dir = "";

        xPos = window.event.offsetX; 
        offset = 8;
 
        if (xPos<offset) dir += "w";
        else if (xPos > el.offsetWidth-offset) dir += "e";

        return dir;
}

function doDown() {
        var el = getReal(event.srcElement, "className", "resizeMe");
 				var el2 = getReal(event.srcElement, "className", "resizeMe2");
 
        if (el == null) {
                theobject = null;
                theobject2 = null;
                return;
        }                

        dir = getDirection(el);
        if (dir == "") return;

        theobject = new resizeObject();
                
        theobject.el = el;
        theobject.dir = dir;

        theobject.grabx = window.event.clientX; 
        theobject.width = el.offsetWidth; 
        theobject.left = el.offsetLeft; 

				theobject2 = new resizeObject();
		 		theobject2.el2 = el2;
      
       theobject2.width = el2.offsetWidth; 
       theobject2.left = el2.offsetLeft; 
        
        window.event.returnValue = false;
        window.event.cancelBubble = true;
}

function doUp() {
        if (theobject != null) {
                theobject = null;
                theobject2 = null;
        }
}

function doMove() {
        var el, xPos, str, xMin ;
        xMin = 0;  

        el = getReal(event.srcElement, "className", "resizeMe");
		var el2 = getReal(event.srcElement, "className", "resizeMe2");
		
		
        if (el.className == "resizeMe") {
                str = getDirection(el);
                if (str == "") str = "default";
                else str += "-resize";
                el.style.cursor = str;
        }
        
        if(theobject != null) {
                if (dir.indexOf("e") != -1){
                        theobject.el.style.width = Math.max(xMin, theobject.width + window.event.clientX - theobject.grabx);  
                      //  theobject2.el2.style.left =  Math.min(theobject.left + window.event.clientX - theobject.grabx, theobject2.left-theobject.grabx);
                       // theobject2.el2.style.width =  Math.max(theobject.width - window.event.clientX + theobject.grabx, theobject2.width + theobject.grabx); 
                        
                        //theobject2.el2.style.left = Math.max(theobject.width + window.event.clientX - theobject.grabx, theobject2.left+theobject2.width);
         				}
                if (dir.indexOf("w") != -1) {
                        theobject.el.style.left = Math.min(theobject.left + window.event.clientX - theobject.grabx, theobject.left + theobject.width - xMin);
                        theobject.el.style.width = Math.max(xMin, theobject.width - window.event.clientX + theobject.grabx); 
                } 
                window.event.returnValue = false;
                window.event.cancelBubble = true;
        } 
}


function getReal(el, type, value) {
        temp = el;
        while ((temp != null) && (temp.tagName != "BODY")) {
                if (eval("temp." + type) == value) {
                        el = temp;
                        return el;
                }
                temp = temp.parentElement;
        }
        return el;
}

document.onmousedown = doDown;
document.onmouseup   = doUp;
document.onmousemove = doMove;

 