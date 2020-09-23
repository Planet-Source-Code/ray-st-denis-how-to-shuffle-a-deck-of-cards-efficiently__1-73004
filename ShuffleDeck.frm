VERSION 5.00
Begin VB.Form frmShuffleDeck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ShuffleDeck - Ray St.Denis"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   9855
   End
   Begin VB.CommandButton cmdShuffleDeck 
      Caption         =   "Shuffle Deck"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmShuffleDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------------
' Author:    Ray St.Denis
' Contact:   stdenis@webtech.on.ca
' Date:      March 17, 2010
' Description:
'   This example is to show how to shuffle a deck of cards using an extremely fast
'   algorithm that does not have to check if a card has already been chosen.
'   What this means is that the code is very fast making it very scalable as well.
'   Although the example uses a deck of cards, it could just as easily be used
'   to shuffle any series of objects with thousands of items.
'------------------------------------------------------------------------------------

Dim DeckSize As Long       'Contains the number of cards in deck
Dim Deck() As String       'Contains the visual card names in the deck
Dim ShuffledDeck() As Long 'Contains the shuffled index to the card deck

Private Sub Form_Load()
   Randomize Timer         'Start Randomizing Engine with a unique seed
   InitializeDeck          'Load deck into memory
   DisplayCurrentDeck      'Display current state of deck
End Sub

Private Sub cmdShuffleDeck_Click()
   ShuffledDeck = ShuffleDeck(DeckSize)
   DisplayCurrentDeck      'Display Current state of deck
End Sub

Private Sub DisplayCurrentDeck()
   Dim Result As String, I As Long, J As Long, Index As Long
   Result = ""
   Index = 0
   For I = 1 To 4
       For J = 1 To 13
           Index = Index + 1
           Result = Result & Deck(ShuffledDeck(Index)) & vbTab
       Next J
       Result = Result & vbCrLf
   Next I
   txtResult.Text = Result
End Sub

Private Sub InitializeDeck()
   Dim Suit As Long, Face As Long, Card As Long
   DeckSize = 52 'Use 52 card deck
   ReDim Deck(1 To DeckSize)
   'Load up the visual card names of the deck
   Card = 0
   For Suit = 1 To 4
       For Face = 2 To 14
           Card = Card + 1
           Select Case Face
                  Case 2 To 10
                       Deck(Card) = CStr(Face)
                  Case 11
                       Deck(Card) = "J"
                  Case 12
                       Deck(Card) = "Q"
                  Case 13
                       Deck(Card) = "K"
                  Case 14
                       Deck(Card) = "A"
           End Select
           Select Case Suit
                  Case 1
                       Deck(Card) = Deck(Card) & " of D"
                  Case 2
                       Deck(Card) = Deck(Card) & " of H"
                  Case 3
                       Deck(Card) = Deck(Card) & " of S"
                  Case 4
                       Deck(Card) = Deck(Card) & " of C"
           End Select
       Next Face
   Next Suit
   'Initialize Unshuffled Deck
   ReDim ShuffledDeck(DeckSize)
   For Card = 1 To 52
       ShuffledDeck(Card) = Card
   Next Card
End Sub

Private Function Pick(ByVal Min As Long, ByVal Max As Long) As Long
    'Pick a random integer between Min and Max inclusive
    Pick = Int((Max - Min + 1) * Rnd + Min)
End Function

Private Function ShuffleDeck(ByVal Max As Long) As Variant
    'One pass shuffle of entire deck
    'Very fast and scalable since it does not need to check for duplicates
    Dim N As Long, I As Long
    Dim Original() As Long, Index() As Long
    ReDim Original(1 To Max), Index(1 To Max)
    For I = 1 To Max
        Original(I) = I
    Next I
    For I = Max To 1 Step -1
        N = Pick(1, I)
        Index(I) = Original(N)
        If N <> I Then
           Original(N) = Original(I)
        End If
    Next I
    ShuffleDeck = Index
End Function

