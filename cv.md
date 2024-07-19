![](https://sun9-27.userapi.com/impf/c9552/v9552204/201a/8NZGR621EbI.jpg?quality=96&as=32x32,48x48,72x72,108x108,160x160&sign=2541bb21f58b00d3f52ab394c81028b2&u=HnRzfEnVSPyYNNhB2lckvioHZ2kVl2gSP_qas8bwrKM&cs=200x200)
# Alexander Bogonosov
**Contact information:**  
  
**Location:** *Vitebsk, Belarus*  
**e-mail:** *maximus999@tut.by*  
**Telegram:** *@Maximus9991982*  
- - -  
**About me:**  
  
*I'm 42 years old. I work at Beltelecom as an engineer since 2005*  
*My goal is to make my life better.*  
*Want to learn new.*  
*Diligent. Responsible.*  
*Less words, more action.*  
- - -  
**Education:**  
  
*Belorussian Academy of communications 1999 - 2005 (specialization: telecommunication network software, studied: C, C++, HTML, JS, PHP, SQL, Pascal, VB)*  
- - -  
**Courses:**  
 
* *JS/FE Pre-school 2024 Q2 (in progress)*

  
- - -  
**My projects:**  
 
* *[My CV](https://maximus9991982.github.io/rsschool-cv/cv)*

  
- - -  
**Code example:**  
 
*This is my last VBA-macro, which makes my routine job easer:*    
``` VBA
Sub PenaltyCalculationAndDate()  '' макрос для автоматического расчета, вводим вручную только первые 2 месяца
If MsgBox("Произвести расчет?", vbQuestion + vbYesNo) = 6 Then

Dim isEmptyGraph As Boolean
Dim GraphColumn As Integer

Application.ScreenUpdating = False

isEmptyGraph = False
valueColumn = 14

Range("O5:AC5").Select
Selection.ClearContents
Range("M6:AC6").Select
Selection.ClearContents
Range("M5:M5").Select
    
'' Заполняем график датами иэ эталонных значений
Do
Cells(5, valueColumn + 1) = Cells(9, valueColumn + 1)
valueColumn = valueColumn + 1
Loop While Cells(5, valueColumn) <= Cells(3, 15)
'' -----------------------------------------------------------------------------------

'' проверка на выходные
valueColumn = 13
Do
If Weekday(Cells(5, valueColumn), vbMonday) = 6 Then
Cells(5, valueColumn) = Cells(5, valueColumn) - 1
Else
If Weekday(Cells(5, valueColumn), vbMonday) = 7 Then
Cells(5, valueColumn) = Cells(5, valueColumn) - 2
End If
End If
valueColumn = valueColumn + 1
Loop While Cells(5, valueColumn) <= Cells(3, 15)
'' -----------------------------------------------------------------------------------

valueColumn = 13
Do        
Cells(4, 6) = Cells(5, valueColumn)    '' копируем дату из графика в расчет
Cells(6, valueColumn) = Cells(24, 10)  '' копируем значение штрафа в график
valueColumn = valueColumn + 1
        
If Cells(5, valueColumn) = Empty Then  '' проверяем закончились ли даты в графике
isEmptyGraph = True
End If
            
Loop While (Not isEmptyGraph)
End If

Application.ScreenUpdating = True
End Sub
```  
- - -  
**Languages:**  
  
* *English (Upper-Intermediate (B2), according to online test: [Puzzle English](https://puzzle-english.com) )*
![](https://sun9-53.userapi.com/impg/sFJuQHHOy-9dHbVdq1PCmHgPA5dXdR4xuvCy_g/dqxXj4vdAjo.jpg?size=1211x375&quality=95&sign=9ef373885ae1e5117e85c33bce05f93b&type=album)

* *Belarussian (native)*  
* *Russian (native)*
