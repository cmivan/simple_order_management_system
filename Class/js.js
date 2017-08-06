//js 容错模式
function killErrors() {
return true;
}
window.onerror =killErrors;


//防止刷新的
function document.onkeydown()       
{ 
if (event.keyCode==116){
event.keyCode = 0;
return false;
}}



if (window.Event) document.captureEvents(Event.MOUSEUP);
function nocontextmenu() 
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
function norightclick(e) 
{
if (window.Event) 
{
if (e.which == 2 || e.which == 3)
return false;
}
else
if (event.button == 2 || event.button == 3)
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
}
document.oncontextmenu = nocontextmenu; //对ie5.0以上
document.onmousedown = norightclick; //对其它浏览器
