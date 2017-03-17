System.Gadget.onDock=Dock;
System.Gadget.onUndock=unDock;
window.onload = System.Gadget.docked ? Dock : unDock;

////////////////////////////////////////////////////////////////////////////////
//
// Docked
//
////////////////////////////////////////////////////////////////////////////////
function Dock(){    
    with(document.body.style)
        width=130,
        height=230;

	control.setDock();

	with(headerImg.style)
		position="absolute",
		top=0, 
		left=0,
		width=130,
		height=30;

	with(control.style)
		position="absolute",
		top=30,
		left=0,
		width=130,
		height=200;
};
////////////////////////////////////////////////////////////////////////////////
//
//  UnDock
//
////////////////////////////////////////////////////////////////////////////////
function unDock()
{
    with(document.body.style)
       width=800,
       height=600;

	control.setUndock();

	with(headerImg.style)
		position="absolute",
		top=0,
		left=0,
		width=130,
		height=30;

	with(control.style)
		position="absolute",
		top=30,
		left=0,
		width=800,
		height=570;
};