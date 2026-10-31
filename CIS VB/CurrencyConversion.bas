Attribute VB_Name = "CurrencyConversion"
Option Explicit
Public tmpDollar As String
Public tmpCents As String

    Function ConvertCurrencyToEnglish(ByVal MyNumber, tmpForm As String)
         Dim temp
         Dim Dollars, Cents
         Dim DecimalPlace, Count
         Dim tmpMeasure As String
         ReDim Place(9) As String
         Place(2) = " Thousand "
         Place(3) = " Million "
         Place(4) = " Billion "
         Place(5) = " Trillion "

         ' Convert MyNumber to a string, trimming extra spaces.
         MyNumber = Trim(str(MyNumber))

         ' Find decimal place.
         DecimalPlace = InStr(MyNumber, ".")

         ' If we find decimal place...
         If DecimalPlace > 0 Then
            ' Convert cents
            temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Cents = ConvertTens(temp)

            ' Strip off cents from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
         End If

         Count = 1
         Do While MyNumber <> ""
            ' Convert last 3 digits of MyNumber to English dollars.
            temp = ConvertHundreds(Right(MyNumber, 3))
            If temp <> "" Then Dollars = temp & Place(Count) & Dollars
            If Len(MyNumber) > 3 Then
               ' Remove last 3 converted digits from MyNumber.
               MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
               MyNumber = ""
            End If
            Count = Count + 1
         Loop

         ' Clean up dollars.
         Select Case Dollars
            Case ""
               Dollars = ""
            Case "One"
               Dollars = "One Dollar"
            Case Else
                'Dollars = Dollars & " And"
                Dollars = Dollars
         End Select

         ' Clean up cents.
         Select Case Cents
            Case ""
               Cents = ""
            Case "One"
               Cents = " And  Cent One"
            Case Else
               Cents = " And " & "Cents " & Cents
         End Select
        
         ConvertCurrencyToEnglish = Dollars & Cents
         tmpMeasure = ConvertCurrencyToEnglish
         
         'FrmPayMember.txttmpDollar = ConvertCurrencyToEnglish
        'If Len(FrmPayMember.txttmpDollar) > 55 Then
            If tmpForm = "MedixPayCLN" Then
                frmPaymentCLNInvoice.txttmpDollar = Dollars
                frmPaymentCLNInvoice.txttmpCents = Cents
            ElseIf tmpForm = "MedixPayMBM" Then
                frmPaymentMemberAuto08.txttmpDollar = Dollars
                frmPaymentMemberAuto08.txttmpCents = Cents
            End If
        'End If
      End Function

      Private Function ConvertHundreds(ByVal MyNumber)
         Dim Result As String

         ' Exit if there is nothing to convert.
         If Val(MyNumber) = 0 Then Exit Function

         ' Append leading zeros to number.
         MyNumber = Right("000" & MyNumber, 3)

         ' Do we have a hundreds place digit to convert?
         If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
         End If

         ' Do we have a tens place digit to convert?
         If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & " And " & ConvertTens(Mid(MyNumber, 2))
         Else
            ' If not, then convert the ones place digit.
            'Result = Result & " And "
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
         End If

         ConvertHundreds = Trim(Result)
      End Function

      Private Function ConvertTens(ByVal MyTens)
         Dim Result As String

         ' Is value between 10 and 19?
         If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
               Case 10: Result = "Ten"
               Case 11: Result = "Eleven"
               Case 12: Result = "Twelve"
               Case 13: Result = "Thirteen"
               Case 14: Result = "Fourteen"
               Case 15: Result = "Fifteen"
               Case 16: Result = "Sixteen"
               Case 17: Result = "Seventeen"
               Case 18: Result = "Eighteen"
               Case 19: Result = "Nineteen"
               Case Else
            End Select
         Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
               Case 2: Result = "Twenty "
               Case 3: Result = "Thirty "
               Case 4: Result = "Forty "
               Case 5: Result = "Fifty "
               Case 6: Result = "Sixty "
               Case 7: Result = "Seventy "
               Case 8: Result = "Eighty "
               Case 9: Result = "Ninety "
               Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
         End If

         ConvertTens = Result
      End Function

      Private Function ConvertDigit(ByVal MyDigit)
         Select Case Val(MyDigit)
            Case 1: ConvertDigit = "One"
            Case 2: ConvertDigit = "Two"
            Case 3: ConvertDigit = "Three"
            Case 4: ConvertDigit = "Four"
            Case 5: ConvertDigit = "Five"
            Case 6: ConvertDigit = "Six"
            Case 7: ConvertDigit = "Seven"
            Case 8: ConvertDigit = "Eight"
            Case 9: ConvertDigit = "Nine"
            Case Else: ConvertDigit = ""
         End Select
      End Function

