Public Class frmChapter5ExercisesSet2
    'Program: Chapter 5 Exercises Set 2
    'Programmer: Christy Mims
    'Date: 3/22/17
    'Description:  This program allows the user to check problems: Exercises 34, 36, 38 on pages 211-212 and Exercises
    '10 and 12 on pages 222-223 and Exercise 2 on page 237.  This program allows the user to do this through the use of
    'buttons, text boxes, list boxes, and labels.  It also uses Subrouteans and Functions to compute the data for the user.

    Private Sub btnComputeTip_Click(sender As Object, e As EventArgs) Handles btnComputeTip.Click
        'This button allows the user to compute their tip.
        Dim Occupation As String
        Dim Bill_Amount, Tip_Percentage, computed_tip As Decimal
        InputInformation(Occupation, Bill_Amount, Tip_Percentage)
        DisplayTip(Occupation, Bill_Amount, Tip_Percentage, computed_tip)
    End Sub
    Sub InputInformation(ByRef Occupation As String, ByRef Bill_Amount As Decimal, ByRef Tip_Percentage As Decimal)
        Occupation = txtOccupation.Text
        Bill_Amount = CDec(txtAmountofBill.Text)
        Tip_Percentage = CDec(txtPercentageTip.Text)
    End Sub
    Sub DisplayTip(Occupation As String, Bill_Amount As Decimal, Tip_Percentage As Decimal, ByRef computed_tip As Decimal)
        computed_tip = (Bill_Amount * (Tip_Percentage / 100))
        txtComputedTip.Text = "Tip the " & Occupation & " " & computed_tip.ToString("C")
    End Sub



    Private Sub btnSemesterGrade_Click(sender As Object, e As EventArgs) Handles btnSemesterGrade.Click
        'This button allows the user to calculate their semester grade.
        Dim NameofPerson As String = ""
        Dim First_Grade, Second_Grade, Third_Grade, Value As Double
        GradeInformation(NameofPerson, First_Grade, Second_Grade, Third_Grade)
        Letter_Grade(NameofPerson, First_Grade, Second_Grade, Third_Grade)
    End Sub
    Sub GradeInformation(ByRef NameofPerson As String, ByRef First_Grade As Double, ByRef Second_Grade As Double,
                         ByRef Third_Grade As Double)
        NameofPerson = txtName.Text
        First_Grade = CDbl(txtFirstGrade.Text)
        Second_Grade = CDbl(txtSecondGrade.Text)
        Third_Grade = CDbl(txtThirdGrade.Text)
    End Sub
    Sub Letter_Grade(NameofPerson As String, First_Grade As Double, Second_Grade As Double, Third_Grade As Double)
        Dim Final_Grade As Double
        Final_Grade = Grade(First_Grade, Second_Grade, Third_Grade)
        Select Case Final_Grade
            Case 90 To 100
                txtFinalGrade.Text = NameofPerson & ": A"
            Case 80 To 89
                txtFinalGrade.Text = NameofPerson & ": B"
            Case 70 To 79
                txtFinalGrade.Text = NameofPerson & ": C"
            Case 60 To 69
                txtFinalGrade.Text = NameofPerson & ": D"
            Case Is <= 59
                txtFinalGrade.Text = NameofPerson & ": F"
        End Select

    End Sub
    Function Grade(First_Grade As Double, Second_Grade As Double, Third_Grade As Double) As Decimal
        Dim value As Double
        If First_Grade <= Second_Grade & Third_Grade Then
            value = (Third_Grade + Second_Grade) / 2
        End If
        If First_Grade & Third_Grade >= Second_Grade Then
            value = (First_Grade + Third_Grade) / 2
        End If
        If Second_Grade & First_Grade >= Third_Grade Then
            value = (First_Grade + Second_Grade) / 2
        End If
        If First_Grade = Second_Grade = Third_Grade Then
            value = value
        End If
        Return value
    End Function





    Private Sub btnBirthdaySong_Click(sender As Object, e As EventArgs) Handles btnBirthdaySong.Click
        'This button allows the user to enter a name and then display the Happy Birthday Song with a name.
        Dim birthday_person As String
        birthday_person = CStr(txtFirstName.Text)
        Verse_3(birthday_person)
    End Sub
    Sub Verse_3(birthday_person As String)
        lstSong.Items.Add("Happy Birthday to you!")
        lstSong.Items.Add("Happy Birthday to you!")
        lstSong.Items.Add("Happy Birthday, dear " & birthday_person & ".")
        lstSong.Items.Add("Happy Birthday to you!")
    End Sub



    Private Sub btnCalculateNew_Click(sender As Object, e As EventArgs) Handles btnCalculateNew.Click
        'This button allows the user to determine the minimum payment for their credit card.
        Dim oldBalance, charges, credits, newBalance, minPayment As Decimal
        InputData(oldBalance, charges, credits)
        CalculateNewValues(oldBalance, charges, credits, newBalance, minPayment)
        DisplayData(newBalance, minPayment)
    End Sub
    Sub InputData(ByRef oldBalance As Decimal, ByRef charges As Decimal, ByRef credits As Decimal)
        oldBalance = CDec(txtOldBalance.Text)
        charges = CDec(txtCharges.Text)
        credits = CDec(txtCredits.Text)
    End Sub
    Sub CalculateNewValues(oldBalance As Decimal, charges As Decimal, credits As Decimal, ByRef newBalance As Decimal,
                          ByRef minPayment As Decimal)
        Dim finance_charge As Decimal
        finance_charge = (0.015D * oldBalance)
        newBalance = (oldBalance + charges + finance_charge) - credits
        If newBalance > 20D Then
            minPayment = 20D + (0.1D * (newBalance - 20D))
        ElseIf newBalance <= 20D Then
            minPayment = newBalance
        End If
    End Sub
    Sub DisplayData(ByRef newBalance As Decimal, ByRef minPayment As Decimal)
        txtNewBalance.Text = newBalance.ToString("C")
        txtMinPayment.Text = minPayment.ToString("C")
    End Sub




    Private Sub btnCalculateOvertime_Click(sender As Object, e As EventArgs) Handles btnCalculateOvertime.Click
        'This button calculates the weekly pay for the user.
        Dim hours, payPerHour, overtimeHours, pay As Decimal
        Input_Data(hours, payPerHour)
        CalculateValues(hours, payPerHour, overtimeHours, pay)
        DisplayNewData(overtimeHours, pay)
    End Sub
    Sub Input_Data(ByRef hours As Decimal, ByRef payPerHour As Decimal)
        hours = CDec(txtHours.Text)
        payPerHour = CDec(txtpayPerHour.Text)
    End Sub
    Sub CalculateValues(hours As Decimal, payPerHour As Decimal, ByRef overtimeHours As Decimal, ByRef pay As Decimal)
        If hours > 40 Then
            overtimeHours = hours - 40
            pay = (40 * payPerHour) + ((payPerHour / 2) + payPerHour) * overtimeHours
        Else
            overtimeHours = 0
            pay = hours * payPerHour
        End If
    End Sub
    Sub DisplayNewData(ovetimeHours As Decimal, pay As Decimal)
        txtOvertimeHours.Text = ovetimeHours
        txtPay.Text = pay.ToString("C")
    End Sub




    Private Sub btnComputeCost_Click(sender As Object, e As EventArgs) Handles btnComputeCost.Click
        'This button computes the cost of food and drinks for the user and displays how many they ordered and how much it costs.
        Dim pizza_slices, fries, drinks, total As Decimal
        FoodOrderInfo(pizza_slices, fries, drinks)
        DisplayTotals(pizza_slices, fries, drinks, total)
    End Sub
    Sub FoodOrderInfo(ByRef pizza_slices As Decimal, ByRef fries As Decimal, ByRef drinks As Decimal)
        pizza_slices = CDec(txtPizzaSlices.Text)
        fries = CDec(txtFries.Text)
        drinks = CDec(txtDrinks.Text)
    End Sub
    Sub DisplayTotals(pizza As Decimal, fries As Decimal, drinks As Decimal, total As Decimal)
        Dim foodCost, Pizza_Price, Frie_Price, Drink_Price As Decimal
        foodCost = foodTotals(total, pizza, fries, drinks)
        Pizza_Price = pizzatotals(pizza)
        Frie_Price = frenchfries(fries)
        Drink_Price = Drinkprice(drinks)
        Dim format_string As String = " {0, -13} {1, 9} {2, 12:C2} "
        lstResults.Items.Add(String.Format(format_string, "ITEM", "QUANTITY", "PRICE"))
        lstResults.Items.Add(String.Format(format_string, "pizza slices", pizza, Pizza_Price))
        lstResults.Items.Add(String.Format(format_string, "fries", fries, Frie_Price))
        lstResults.Items.Add(String.Format(format_string, "soft drinks", drinks, Drink_Price))
        lstResults.Items.Add(String.Format(format_string, "TOTAL", " ", foodCost))
    End Sub
    Function pizzatotals(ByRef pizza As Decimal) As Decimal
        Return 1.75D * pizza
    End Function
    Function frenchfries(ByRef fries As Decimal) As Decimal
        Return 2D * fries
    End Function
    Function Drinkprice(ByRef drinks As Decimal) As Decimal
        Return 1.25D * drinks
    End Function
    Function foodTotals(ByRef total As Decimal, pizza As Decimal, fries As Decimal,
                        drinks As Decimal) As Decimal
        Dim Pizza_Cost, Frie_Cost, Drink_Cost As Decimal
        Pizza_Cost = pizzatotals(pizza)
        Frie_Cost = frenchfries(fries)
        Drink_Cost = Drinkprice(drinks)
        total = Pizza_Cost + Frie_Cost + Drink_Cost
        Return total
    End Function
End Class
