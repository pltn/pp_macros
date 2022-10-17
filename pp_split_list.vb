Sub splitList()
    ''' Текущий текст  
    Set tr = ActiveWindow.Selection.TextRange2
    
    ''' Текущий слайд
    Set s = ActiveWindow.View.Slide
    
    ''' Длина списка
    ps = tr.Paragraphs.Count
    
    ''' Положение списка X
    x = ActiveWindow.Selection.ShapeRange(1).Left
    
    ''' Положение списка Y
    y = ActiveWindow.Selection.ShapeRange(1).Top
  
    ''' Удивительный костыль: без символа перевода строки VBA не считает последнюю строку элементом списка
    tr.InsertAfter (vbCr)
  
	''' Для каждого элемента списка
    For i = 1 To ps
        With tr.Paragraphs.Item(i)
            ''' Создаём надпись
            Set tmpTextBox = s.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=10, Top:=y + 50 * (i - 1), Width:=200, Height:=50)
            ''' Копируем параграф из исходной надписи
            tr.Paragraphs.Item(i).Copy
            ''' Вставляем в созданную надпись
            tmpTextBox.TextFrame2.TextRange.Paste
            ''' Если у нас нумерованный список, то
            If tmpTextBox.TextFrame2.TextRange.ParagraphFormat.Bullet.Type = ppBulletNumbered Then
                ''' Выставляем значение начала списка равное номеру элемента в списке
                tmpTextBox.TextFrame2.TextRange.ParagraphFormat.Bullet.StartValue = i
            End If
        End With
    Next
	
    ''' Удаляем костыль
    tr.Text = Mid(tr.Text, 1, Len(tr.Text) - 1)

End Sub