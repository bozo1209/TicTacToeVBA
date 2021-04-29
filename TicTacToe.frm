VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TicTacToe 
   Caption         =   "TicTacToe"
   ClientHeight    =   5748
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8448
   OleObjectBlob   =   "TicTacToe.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j As Integer
Dim isPlayer1 As Boolean
Dim gameBoard(0 To 2, 0 To 2) As String

Sub ticTacToeStart()
    ticTacToe.Show
End Sub

Private Sub button00_Click()
    If isPlayer1 Then
        button00.Caption = "X"
        gameBoard(0, 0) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button00.Caption = "O"
        gameBoard(0, 0) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button00.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button01_Click()
    If isPlayer1 Then
        button01.Caption = "X"
        gameBoard(0, 1) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button01.Caption = "O"
        gameBoard(0, 1) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button01.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button02_Click()
    If isPlayer1 Then
        button02.Caption = "X"
        gameBoard(0, 2) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button02.Caption = "O"
        gameBoard(0, 2) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button02.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button10_Click()
    If isPlayer1 Then
        button10.Caption = "X"
        gameBoard(1, 0) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button10.Caption = "O"
        gameBoard(1, 0) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button10.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button11_Click()
    If isPlayer1 Then
        button11.Caption = "X"
        gameBoard(1, 1) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button11.Caption = "O"
        gameBoard(1, 1) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button11.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button12_Click()
    If isPlayer1 Then
        button12.Caption = "X"
        gameBoard(1, 2) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button12.Caption = "O"
        gameBoard(1, 2) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button12.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button20_Click()
    If isPlayer1 Then
        button20.Caption = "X"
        gameBoard(2, 0) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button20.Caption = "O"
        gameBoard(2, 0) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button20.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button21_Click()
    If isPlayer1 Then
        button21.Caption = "X"
        gameBoard(2, 1) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button21.Caption = "O"
        gameBoard(2, 1) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button21.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub button22_Click()
    If isPlayer1 Then
        button22.Caption = "X"
        gameBoard(2, 2) = "X"
        isPlayer1 = False
        gameStatus.Caption = "player 2 turn"
    Else
        button22.Caption = "O"
        gameBoard(2, 2) = "O"
        isPlayer1 = True
        gameStatus.Caption = "player 1 turn"
    End If
    
    button22.Enabled = False
    
    Call isGameOver
    
End Sub

Private Sub UserForm_Initialize()
    titleLabel.Font.Size = 30
    gameStatus.Font.Size = 20
    isPlayer1 = True
    button00.Font.Size = 30
    button01.Font.Size = 30
    button02.Font.Size = 30
    button10.Font.Size = 30
    button11.Font.Size = 30
    button12.Font.Size = 30
    button20.Font.Size = 30
    button21.Font.Size = 30
    button22.Font.Size = 30
End Sub

Private Sub isGameOver()
    For i = 0 To UBound(gameBoard, 1)
        For j = 0 To UBound(gameBoard, 1)
            If gameBoard(i, 0) = "X" And gameBoard(i, 1) = "X" And gameBoard(i, 2) = "X" Then
                Call gameOver("X")
                GoTo nestedLoopExit
            ElseIf gameBoard(0, j) = "X" And gameBoard(1, j) = "X" And gameBoard(2, j) = "X" Then
                Call gameOver("X")
                GoTo nestedLoopExit
            ElseIf gameBoard(0, 0) = "X" And gameBoard(1, 1) = "X" And gameBoard(2, 2) = "X" Then
                Call gameOver("X")
                GoTo nestedLoopExit
            ElseIf gameBoard(2, 0) = "X" And gameBoard(1, 1) = "X" And gameBoard(0, 2) = "X" Then
                Call gameOver("X")
                GoTo nestedLoopExit
            ElseIf gameBoard(i, 0) = "O" And gameBoard(i, 1) = "O" And gameBoard(i, 2) = "O" Then
                Call gameOver("O")
                GoTo nestedLoopExit
            ElseIf gameBoard(0, j) = "O" And gameBoard(1, j) = "O" And gameBoard(2, j) = "O" Then
                Call gameOver("O")
                GoTo nestedLoopExit
            ElseIf gameBoard(0, 0) = "O" And gameBoard(1, 1) = "O" And gameBoard(2, 2) = "O" Then
                Call gameOver("O")
                GoTo nestedLoopExit
            ElseIf gameBoard(2, 0) = "O" And gameBoard(1, 1) = "O" And gameBoard(0, 2) = "O" Then
                Call gameOver("O")
                GoTo nestedLoopExit
            End If
        Next j
    Next i
    
nestedLoopExit:
    
End Sub

Private Sub gameOver(result As String)
    If result = "X" Then
        gameStatus.Caption = "player 1 win"
    ElseIf result = "O" Then
        gameStatus.Caption = "player 2 win"
    End If
End Sub
