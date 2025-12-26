VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptionsTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Purpose: This class represents an options trade with all its parameters

Option Explicit

Implements INameProvider

Private mstrDirection As String ' "BUY" or "SELL"
Private mintQuantity As Integer
Private mstrOptionType As String ' "PUT" or "CALL"
Private mstrUnderlying As String ' e.g., "QQQ"
Private mstrExpiration As String ' e.g., "16 of DEC" or "12/16/2024"
Private mdblStrike As Double
Private mdblTarget As Double ' Optional target price
Private mdblStopLoss As Double ' Optional stop loss price
Private mdblEntryPrice As Double ' The price at which the option was entered (for PNL calculation)
Private mdblCurrentUnderlyingPrice As Double ' Current underlying price at entry

Private Function INameProvider_GetClassName() As String
    INameProvider_GetClassName = "clsOptionsTrade"
End Function

Property Get Direction() As String
    Direction = mstrDirection
End Property

Property Let Direction(ByVal value As String)
    mstrDirection = UCase(value)
End Property

Property Get Quantity() As Integer
    Quantity = mintQuantity
End Property

Property Let Quantity(ByVal value As Integer)
    mintQuantity = value
End Property

Property Get OptionType() As String
    OptionType = mstrOptionType
End Property

Property Let OptionType(ByVal value As String)
    mstrOptionType = UCase(value)
End Property

Property Get Underlying() As String
    Underlying = mstrUnderlying
End Property

Property Let Underlying(ByVal value As String)
    mstrUnderlying = UCase(value)
End Property

Property Get Expiration() As String
    Expiration = mstrExpiration
End Property

Property Let Expiration(ByVal value As String)
    mstrExpiration = value
End Property

Property Get Strike() As Double
    Strike = mdblStrike
End Property

Property Let Strike(ByVal value As Double)
    mdblStrike = value
End Property

Property Get Target() As Double
    Target = mdblTarget
End Property

Property Let Target(ByVal value As Double)
    mdblTarget = value
End Property

Property Get StopLoss() As Double
    StopLoss = mdblStopLoss
End Property

Property Let StopLoss(ByVal value As Double)
    mdblStopLoss = value
End Property

Property Get EntryPrice() As Double
    EntryPrice = mdblEntryPrice
End Property

Property Let EntryPrice(ByVal value As Double)
    mdblEntryPrice = value
End Property

Property Get CurrentUnderlyingPrice() As Double
    CurrentUnderlyingPrice = mdblCurrentUnderlyingPrice
End Property

Property Let CurrentUnderlyingPrice(ByVal value As Double)
    mdblCurrentUnderlyingPrice = value
End Property

Private Sub Class_Initialize()
    mstrDirection = ""
    mintQuantity = 0
    mstrOptionType = ""
    mstrUnderlying = ""
    mstrExpiration = ""
    mdblStrike = 0
    mdblTarget = 0
    mdblStopLoss = 0
    mdblEntryPrice = 0
    mdblCurrentUnderlyingPrice = 0
End Sub

