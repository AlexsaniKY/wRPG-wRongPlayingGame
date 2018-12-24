Set objExplorer = CreateObject("InternetExplorer.Application")
Set objShell = WScript.CreateObject("WScript.Shell")

'tile definitions
line_arr = Array(_
9723,9596,9597,9487,_
9598,9473,9491,9523,_
9599,9495,9475,9507,_
9499,9531,9515,9547 _
)

'seed the grid with values
dim grid(30,30)
for y = 0 to ubound(grid,1)
	randomize
	for x = 0 to ubound(grid,2)
		'fill with 1s and 0s
		grid(y,x) = Int(2*rnd)
	next
next

'inform neighbor squares
for y = 1 to ubound(grid,1)-1
	for x = 1 to ubound(grid,2)-1
		if (grid(y,x) mod 2 = 1) then
			grid(y,x+1) =grid(y,x+1)+ 8
			grid(y+1,x) =grid(y+1,x)+ 16
			grid(y,x-1) =grid(y,x-1)+ 2
			grid(y-1,x) =grid(y-1,x)+ 4
		end if
	next
next

ch = "<table align=""left"" border=""0"" cellpadding=""0"" cellspacing=""0"" table-layout=""fixed""> <tbody>"
i = 0
'glyph lookup
for y = 1 to ubound(grid,1)-1
	ch = ch + "<tr style = ""border:none;"">"
	for x = 1 to ubound(grid,2)-1
		i = grid(y,x)
		if (i mod 2 = 1) then
			i = Int((i-1)/2)
			ch = ch & "<td><p>" & "&#" & line_arr(i) & "</p></td>"
		else
			ch = ch & "<td><p>" & chrw(32) & "</p></td>"
		end if
	next
	ch = ch + "</tr>"
next
ch = ch + "</tbody></table>"

With objExplorer
    .Navigate "about:blank"
    '.Visible = 1
    .Document.Title = "Show Image"
    .Toolbar=False
    .Statusbar=False
    .Top=400
    .Left=400
    .Height=200
    .Width=200
	.Document.head="<!doctype html><meta charset=utf-8><script type= ""text/css"">table{border-spacing: 0;} td, th, tr {text-align: center; overflow:hidden; white-space:nowrap; padding:0; margin:0; border:0; vertical-align: baseline; line-height:.5;} td p{line-height:.5;} table {font-family:monospace;} tr{width:1%;}</script>"
    .Document.Body.InnerHTML = "<div>" & ch & "</div>"
	'"<img src='C:\Users\Jill\Pictures\face.png'>"
End With



'Wscript.sleep 4000
objExplorer.visible = true
objShell.AppActivate objExplorer

objExplorer.document.focus()

'objExplorer.top = 0
'objExplorer.left = 0
'objExplorer.Document.Body.InnerHTML ="goodbye"
'msgbox " "