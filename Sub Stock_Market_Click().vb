Sub Stock_Market_Click()

'String Variables
Dim Year As String
Dim Ticker As String

'Array Variables
Dim ArrVo(800000) As Double
Dim ArrVf(800000) As Double
Dim ArrTicker(800000) As String
Dim ArrVol(800000) As Double
Dim ArrVol1(800000) As Double


'Cont Variables
Dim Contx As Long
Dim ContArrTicker As Long
Dim ContArrVo As Long
Dim ContArrVf As Long
Dim ContArrVol As Long
Dim ContArrVol1 As Long


'Turn off Automatic Calculations to speed up Macro
Application.Calculation = xlCalculationManual


'Clean and SetUp the Analysis
Range("D:O").ClearContents
Range("D:O").Interior.Color = RGB(248, 248, 248)

Range("R2:S4").ClearContents
Range("R2:S4").Interior.Color = RGB(248, 248, 248)

Range("L1").Value = "Ticker"
Range("M1").Value = "Yearly Change"
Range("N1").Value = "Perncent Change"
Range("O1").Value = "Total Stock Volume"
Range("L1:O1").Interior.Color = RGB(0, 0, 102)
Range("L1:O1").Font.Color = RGB(255, 255, 255)

'Get the Value from Range B5
Year = Range("B5").Value

'Turn off Screen Updating to speed up Macro
Application.ScreenUpdating = False


'Conditional validation if year is Null
If Year = "" Then
    MsgBox ("Please select a year")
Else

'Get Info from Source Sheets
    Sheets(Year).Range("A:G").Copy
    Range("D:J").Select
    ActiveSheet.Paste

    MsgBox ("Analysing " + Year + " year")

'Set up firt Conts
    Contx = 1
    ContResX = 2
    Ticker = Cells(Contx, 4).Value
    Ticker1 = Cells(Contx + 1, 4).Value

        Do While Len(Ticker) > 0

'Validation of first row
            If Ticker = "<ticker>" Then

                Contx = Contx + 1
                Ticker = Cells(Contx, 4).Value

            Else

'Set up variables
                ContArrTicker = 0
                ContArrVo = 0
                ContArrVf = 0
                ContArrVol = 0
                ContArrVol1 = 1
                ArrVol1(0) = 0


                While Ticker = Ticker1

'Fill Arrays Info
                    Ticker = Cells(Contx, 4).Value
                    ArrTicker(ContArrTicker) = Cells(Contx, 4).Value
                    ArrVo(ContArrVo) = Cells(Contx, 6).Value
                    ArrVf(ContArrVf) = Cells(Contx, 9).Value
                    ArrVol(ContArrVol) = Cells(Contx, 10).Value
                    ArrVol1(ContArrVol1) = ArrVol(ContArrVol) + ArrVol1(ContArrVol1 - 1)

'Pull results from Arrays
                    Cells(ContResX, 12) = ArrTicker(ContArrTicker)
                    Cells(ContResX, 13) = ArrVf(ContArrVf) - ArrVo(0)

'Cell Validation - Interior Color
                    If Cells(ContResX, 13) >= 0 Then
                        Cells(ContResX, 13).Interior.Color = RGB(0, 204, 0)
                    Else
                        Cells(ContResX, 13).Interior.Color = RGB(255, 0, 0)
                    End If

'Cell Validation - to avoid 0/0 division
                    If (ArrVf(ContArrVf) = 0 And ArrVo(0) = 0) Or ArrVo(0) = 0 Then
                        Cells(ContResX, 14) = 0
                    Else
                        Cells(ContResX, 14) = (ArrVf(ContArrVf) / ArrVo(0)) - 1
                    End If



                    Cells(ContResX, 15) = ArrVol1(ContArrVol1)

'Conts to move forward on the second loop
                    ContArrVf = ContArrVf + 1
                    ContArrVo = ContArrVo + 1
                    ContArrVol = ContArrVol + 1
                    ContArrVol1 = ContArrVol1 + 1

                    Contx = Contx + 1
                    Ticker1 = Cells(Contx, 4).Value

                Wend
'Conts to move forward on the first loop
                    ContResX = ContResX + 1
                    Ticker = Cells(Contx + 1, 4).Value
            End If
        Loop
End If



MsgBox ("Analysis First part is Finished")

'***********************************************************


'Variables
Dim SizeArr As Double
Dim MaxValueChange As Double
Dim MinValueChange As Double
Dim MaxVol As Double
Dim ContBonus As Double




'Functions to find Max and Min
SizeArr = Application.WorksheetFunction.CountA(Range("N:N"))
MaxValueChange = Application.WorksheetFunction.Max(Range("N:N"))
MinValueChange = Application.WorksheetFunction.Min(Range("N:N"))
MaxVol = Application.WorksheetFunction.Max(Range("O:O"))

'Pull Value Data
Cells(2, 19).Value = MaxValueChange
Cells(3, 19).Value = MinValueChange
Cells(4, 19).Value = MaxVol

'Pull Ticker Data
For ContBonus = 1 To SizeArr

If Cells(ContBonus, 14).Value = MaxValueChange Then

   Cells(2, 18).Value = Cells(ContBonus, 14 - 2)
   
   Else
   
End If

Next

'ContBonus = 1

For ContBonus = 1 To SizeArr

If Cells(ContBonus, 14).Value = MinValueChange Then

   Cells(3, 18).Value = Cells(ContBonus, 14 - 2)
   
   Else
   
End If

Next

'ContBonus = 1
For ContBonus = 1 To SizeArr

If Cells(ContBonus, 15).Value = MaxVol Then

   Cells(4, 18).Value = Cells(ContBonus, 15 - 3)
   
   Else
   
End If

Next


'Turn on Automatic Calculations to speed up Macro
Application.Calculation = xlCalculationAutomatic

MsgBox ("Analysis Bonus part is Finished")

End Sub









