Attribute VB_Name = "RateCards"
Option Explicit
Option Base 1

'Module for the purpose of storing and retrieving the position of an employee.
'Stores the rate card table as a manually input array and retrieves the position title from there.
'To find a person's position, all it requires is the date and their pay level.

Function retrieve_rate_card() As Variant
'Will return an array that contains the rate card

    Dim rateCard(7, 2) As Variant
    Dim firstTranche() As Variant
    Dim secondTranche() As Variant
    Dim thirdTranche() As Variant
    
    rateCard(1, 1) = DateValue("2020/07/01")
    rateCard(2, 1) = DateValue("2019/01/01")
    rateCard(3, 1) = DateValue("2017/07/01")
    rateCard(4, 1) = DateValue("2016/07/01")
    rateCard(5, 1) = DateValue("2013/09/01")
    rateCard(6, 1) = DateValue("2011/03/01")
    rateCard(7, 1) = DateValue("2006/07/01")
    
    firstTranche = Array("Partner", "Director", "Senior Manager", "Manager", "Supervisor", "Senior Accountant 1", "Senior Accountant 2", "Intermediate Accountant 1", "Intermediate Accountant 2", "Accountant", "Senior Administration", "Administration")
    
    rateCard(1, 2) = Array(firstTranche, Array(650, 595, 550, 525, 455, 405, 375, 325, 305, 280, 235, 175))
    rateCard(2, 2) = Array(firstTranche, Array(625, 570, 525, 500, 430, 380, 350, 300, 280, 255, 210, 150))
    rateCard(3, 2) = Array(firstTranche, Array(625, 570, 525, 500, 420, 370, 340, 295, 275, 250, 210, 150))
                           
    secondTranche = Array("Partner", "Associate Director", "Senior Manager", "Manager", "Supervisor", "Senior Accountant 1", "Senior Accountant 2", "Intermediate Accountant 1", "Intermediate Accountant 2", "Accountant", "Senior Secretary", "Secretary")
    
    rateCard(4, 2) = Array(secondTranche, Array(595, 540, 500, 475, 400, 350, 325, 280, 260, 240, 200, 140))
    rateCard(5, 2) = Array(secondTranche, Array(565, 495, 440, 410, 355, 300, 275, 205, 175, 145, 175, 130))
    rateCard(6, 2) = Array(secondTranche, Array(525, 450, 400, 375, 325, 275, 250, 190, 160, 135, 160, 120))
    
    thirdTranche = Array("Appointee / Director", "Associate Director", "Supervisor", "Secretary")
    
    rateCard(7, 2) = Array(thirdTranche, Array(435, 300, 185, 130))
    
    retrieve_rate_card = rateCard

End Function

Public Function return_position(ByVal wipDate As Date, ByVal hourlyRate As Currency) As String
'Takes in two parameters, date and rate, and will return the person's position based on this.
'Uses the rate card function above.

    Dim rateCard As Variant
    Dim i As Byte
    Dim j As Byte
    
    rateCard = retrieve_rate_card()
    'This will be the value returned if the position cannot be found
    return_position = "INVALID RATE"
    
    'Iterating through all dates in the rate card, starting with most recent
    For i = 1 To UBound(rateCard)
        
        'If in the time period we are looking for
        If wipDate >= rateCard(i, 1) Then
            
            'Iterating through all rates in this time period
            For j = 1 To UBound(rateCard(i, 2)(1))
                
                'When we find a rate that matches
                If hourlyRate = rateCard(i, 2)(2)(j) Then
                    'Ending the function as soon as we find our result
                    return_position = rateCard(i, 2)(1)(j)
                    Exit Function
                End If
                
            Next j
            
            'This is so if it can't find a matching amount in the time period,
            'we don't want one at all, otherwise the one returned will be wrong.
            Exit Function
            
        End If
        
    Next i

End Function

