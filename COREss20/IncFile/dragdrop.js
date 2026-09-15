
var x=y=o=null;


if (document.layers) {
    window.captureEvents(Event.MOUSEDOWN | Event.MOUSEMOVE | Event.MOUSEUP);
    window.onmousedown=down, window.onmouseup=up, window.onmousemove=move;
}
else if (document.all)
    document.onmousedown=down, document.onmouseup=up, document.onmousemove=move;

function down(e) {
try {

    if ( document.all && window.event.srcElement.parentElement.parentElement.parentElement.parentElement)
     o=window.event.srcElement.parentElement.parentElement.parentElement.parentElement, x=window.event.offsetX, y=window.event.offsetY;
    else {
        if (document.layers) {
            if (document.layers['o'+e.target.name]) {
                o=document.layers['o'+e.target.name], x=e.layerX, y=e.layerY;
            }
            else return true;
        }
   }

   return false;
}
catch(er) {

}
}

function move(e) {

    if (document.all && o)
     {

        o.style.posLeft=window.event.clientX-x, o.style.posTop=window.event.clientY-y;
     }
    else if (document.layers && o)
        o.left=e.pageX-x, o.top=e.pageY-y;
    return false;
}

function up() { o=null; }
