Set objExplorer = CreateObject("InternetExplorer.Application")
Set objShell = WScript.CreateObject("WScript.Shell")

'tile definitions
'right = 1, down = 2, left = 4, up = 8
'null, r, d, rd, 
'l, rl, dl, rdl, 
'u, ru, du, rdu, 
'lu, rlu, rdl, rdlu
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
			'use double the side values to sentinel when
			'it was near a tile without being one
			
			'inform the right square
			grid(y,x+1) =grid(y,x+1)+ 8
			'the down square
			grid(y+1,x) =grid(y+1,x)+ 16
			'left square
			grid(y,x-1) =grid(y,x-1)+ 2
			'up square
			grid(y-1,x) =grid(y-1,x)+ 4
		end if
	next
next

ch = "<table align=""left"" border=""0"" cellpadding=""0"" cellspacing=""0"" > <tbody>"
i = 0
'glyph lookup
for y = 1 to ubound(grid,1)-1
	ch = ch + "<tr>"
	
	for x = 1 to ubound(grid,2)-1
		ch = ch & "<td><div>"
		i = grid(y,x)
		if (i mod 2 = 1) then
			i = Int((i-1)/2)
			ch = ch & "&#" & line_arr(i)
		else
			ch = ch & chrw(32)
		end if
		ch = ch & "</div></td>"
	next
	ch = ch + "</tr>"
next
ch = ch + "</tbody></table>"

ch = "<table align=""left""  cellpadding=""0"" cellspacing=""0"" > <tbody>"
for y = 0 to ubound(grid,1)-1
	ch = ch + "<tr>"
	for x = 0 to ubound(grid,2)-1
		ch = ch & "<td id=""" & y & "," & x & """>"
		if y mod 4 = 0 then
			ch = ch & "&#x" & 2588
		else 
			ch = ch & "&#x" & (2590 + (y mod 4))
		end if
		ch = ch & "</td>"
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
    .Height=600
    .Width=400
	.Document.head.InnerHTML="<!doctype html><meta charset=utf-8>"&_
								"<style>"&_
									"td{text-align: center; line-height:1.12em;}"&_ 
								"</style>"
    .Document.Body.InnerHTML = "<div>" & ch & "</div>"
End With

objExplorer.visible = true

Wscript.sleep 4000
objExplorer.Document.getElementById("5,5").innerhtml = "a"


objShell.AppActivate objExplorer

objExplorer.document.focus()

'objExplorer.top = 0
'objExplorer.left = 0
'objExplorer.Document.Body.InnerHTML ="goodbye"
'msgbox " "