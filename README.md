<div align="center">

## ASP TicTacToe


</div>

### Description

ASP version of TicTacToe. Try to beat the computer at getting 3 Xs in a row.
 
### More Info
 
Code is fully self-sufficient, incorporating both the display and game logic. It's interesting to note that the game was converted to ASP from Javascript.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Owen Cutajar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/owen-cutajar.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__4-13.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/owen-cutajar-asp-tictactoe__4-6364/archive/master.zip)

### API Declarations

Source is PD


### Source Code

```
<%
'--------------------------------------------------------------
' ASP TicTacToe v1.0
' (x) Ugh!! 2000 - 18/10/2000
'
' based on Javascript TicTacToe by
' Maximilian Stocker (maxstocker@reallyusefulcomputing.com)
'
' Written for http://www.only-network.com/games Any comments,
' flames, requests, postcards etc to owen@cutajar.net
'--------------------------------------------------------------
Option Explicit
' -- Set up images to use ---
Const IMGx = "x.jpg"
Const IMGo = "o.jpg"
Const IMGblank = "blank.jpg"
' -- Set up game States ---
Const END_Not_Yet = 0
Const END_You_Win = 1
Const END_Computer_Win = 2
Const END_Tie = 3
' Read in board or Initialise
Dim Gameboard
Dim wl1,wl2,wl3,wl4,wl5,wl6,wl7,wl8
If Session("GameBoard") & "" = "" OR Request("PlayAgain") = "Yes" Then
	PlayAgain
end if
GameBoard = Split(Session("GameBoard"),"_")
function GameState()
	wl1 = GameBoard(0) & GameBoard(1) & GameBoard(2)
	wl2 = GameBoard(0) & GameBoard(3) & GameBoard(6)
	wl3 = GameBoard(0) & GameBoard(4) & GameBoard(8)
	wl4 = GameBoard(1) & GameBoard(4) & GameBoard(7)
	wl5 = GameBoard(3) & GameBoard(4) & GameBoard(5)
	wl6 = GameBoard(6) & GameBoard(7) & GameBoard(8)
	wl7 = GameBoard(2) & GameBoard(5) & GameBoard(8)
	wl8 = GameBoard(6) & GameBoard(4) & GameBoard(2)
	if wl1="XXX" or wl2="XXX" or wl3="XXX" or wl4="XXX" or wl5="XXX" or wl6="XXX" or wl7="XXX" or wl8="XXX" Then
		GameState = END_You_Win
	elseif wl1="OOO" or wl2="OOO" or wl3="OOO" or wl4="OOO" or wl5="OOO" or wl6="OOO" or wl7="OOO" or wl8="OOO" Then
		GameState = END_Computer_Win
	elseif Instr ( wl1 & wl5 & wl6 ,"B" ) = 0 then
		GameState = END_Tie
	else
		GameState = END_Not_Yet
	end if
end function
Function Suggest()
	If wl1 = "XXB" or wl1 = "OOB" Then
		Suggest = 2
	elseif wl1 = "XBX" or wl1 = "OBO" Then
		Suggest = 1
	elseif wl1 = "BXX" or wl1 = "BOO" Then
		Suggest = 0
	elseif wl2 = "XXB" or wl2 = "OOB" Then
		Suggest = 6
	elseif wl2 = "XBX" or wl2 = "OBO" Then
		Suggest = 3
	elseif wl2 = "BXX" or wl2 = "BOO" Then
		Suggest = 0
	elseif wl3 = "XXB" or wl3 = "OOB" Then
		Suggest = 8
	elseif wl3 = "XBX" or wl3 = "OBO" Then
		Suggest = 4
	elseif wl3 = "BXX" or wl3 = "BOO" Then
		Suggest = 0
	elseif wl4 = "XXB" or wl4 = "OOB" Then
		Suggest = 7
	elseif wl4 = "XBX" or wl4 = "OBO" Then
		Suggest = 4
	elseif wl4 = "BXX" or wl4 = "BOO" Then
		Suggest = 1
	elseif wl5 = "XXB" or wl5 = "OOB" Then
		Suggest = 5
	elseif wl5 = "XBX" or wl5 = "OBO" Then
		Suggest = 4
	elseif wl5 = "BXX" or wl5 = "BOO" Then
		Suggest = 3
	elseif wl6 = "XXB" or wl6 = "OOB" Then
		Suggest = 8
	elseif wl6 = "XBX" or wl6 = "OBO" Then
		Suggest = 7
	elseif wl6 = "BXX" or wl6 = "BOO" Then
		Suggest = 6
	elseif wl7 = "XXB" or wl7 = "OOB" Then
		Suggest = 8
	elseif wl7 = "XBX" or wl7 = "OBO" Then
		Suggest = 5
	elseif wl7 = "BXX" or wl7 = "BOO" Then
		Suggest = 2
 	elseif wl8 = "XXB" or wl8 = "OOB" Then
		Suggest = 2
	elseif wl8 = "XBX" or wl8 = "OBO" Then
		Suggest = 4
	elseif wl8 = "BXX" or wl8 = "BOO" Then
		Suggest = 6
	else
		Suggest = -1
	end if
end function
sub yourChoice(Position)
	if Session("State") = "Dead" Then
		ReportEnded
	Else
		If GameBoard(Position) <> "B" Then
			ReportTaken
		else
			GameBoard(Position) = "X"
		end if
 	end if
end sub
sub ReportTaken()
	Response.Write "<H2>That square is already occupied. Please select another square.</H2>"
end sub
sub ReportEnded()
	Response.Write "<H2>The game has already ended. To play a new game click the Play Again button.</H2>"
end sub
sub myChoice()
	Dim NewMove
	NewMove = Suggest()
	While NewMove = -1
		Randomize
		NewMove=int(rnd*9)
		If GameBoard(NewMove) <> "B" Then
			NewMove = -1
		End If
	wend
	GameBoard(NewMove) = "O"
end sub
sub ProcessBoard()
	If Session("State") = "Alive" Then
		Select Case GameState()
			Case END_You_Win
				Response.Write "<H2>You won, congratulations!<H2>"
				Session("you") = Session("you") + 1
				Session("State") = "Dead"
			Case END_Computer_Win
				Response.Write "<H2>Gotcha! I win!</H2>"
				Session("computer") = Session("computer") + 1
				Session("State") = "Dead"
			Case END_Tie
				Response.Write "<H2>We tied.</H2>"
				Session("ties") = Session("ties") + 1
				Session("State") = "Dead"
		end Select
	End If
end sub
sub playAgain()
	Session("GameBoard") = "B_B_B_B_B_B_B_B_B"
	Session("State") = "Alive"
end sub
sub Display(CellNum)
	If GameBoard(CellNum) = "B" Then
		Response.Write "<form action=tictactoe.asp method=post>"
		Response.Write "<input type=hidden name=pressed value=" & CellNum & ">"
		Response.Write "<input type=image src=" & IMGblank & " border=0 height=100 width=100>"
		Response.Write "</form>"
	elseif GameBoard(CellNum) = "O" Then
		Response.Write "<img src=" & IMGo & " border=0 height=100 width=100>"
	elseif GameBoard(CellNum) = "X" Then
		Response.Write "<img src=" & IMGx & " border=0 height=100 width=100>"
	end if
end sub
' Main Code
If Request("Pressed") & "" <> "" Then
	YourChoice(Request("Pressed"))
	ProcessBoard
	If GameState() = END_Not_Yet Then
		myChoice
	End If
	ProcessBoard
	' Save Game State
	Session("GameBoard") = Join(GameBoard,"_")
End If
%>
<HTML>
<HEAD>
</HEAD>
<BODY>
<center>
Welcome to Tic-Tac-Toe! You play as the X's and the computer is the O's. Select the square you want to put your X into by clicking them. You cannot occupy a square that is already occupied. The first player to get three squares in a row wins. Good Luck!!
<form name=game action=tictactoe.asp>
<table border=0>
<tr>
<td>
<table border=1>
<tr height=120>
<td><% Display(0) %></td>
<td><% Display(1) %></td>
<td><% Display(2) %></td>
</tr>
<tr height=120>
<td><% Display(3) %></td>
<td><% Display(4) %></td>
<td><% Display(5) %></td>
</tr>
<tr height=120>
<td><% Display(6) %></td>
<td><% Display(7) %></td>
<td><% Display(8) %></td>
</tr>
</table>
</td>
<td>
<table>
<tr><td><input type=text size=5 name=you value=<%=Session("you")%>></td><td>You</td></tr>
<tr><td><input type=text size=5 name=computer value=<%=Session("computer")%>></td><td>Computer</td></tr>
<tr><td><input type=text size=5 name=ties value=<%=Session("ties")%>></td><td>Ties</td></tr>
</table>
</td></tr>
</table>
<form action=tictactoe.asp>
<input type=hidden name=PlayAgain value=Yes>
<input type=submit value="Play Again">
</form>
</center>
</BODY>
</HTML>
```

