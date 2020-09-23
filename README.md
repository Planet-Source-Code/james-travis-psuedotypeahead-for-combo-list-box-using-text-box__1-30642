<div align="center">

## PsuedoTypeahead for Combo/List box using text box


</div>

### Description

To allow a user to typeahead to a specific item in a list/combo box allowing a little more flexability.
 
### More Info
 
Data entry in text box.

Basically this demonstrate how to loop thru a combo/list box control and compare the typed data to the text value displayed in a the control boxes and select the item. Usefully in large selection lists.

Position and selects the first matching item.

Depending on if the items are sorted the jumping around and selection may look unusual.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Travis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-travis.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-travis-psuedotypeahead-for-combo-list-box-using-text-box__1-30642/archive/master.zip)





### Source Code

```
'You will need 3 controls
'TextBox = Text1
'ComboBox = Combo1
'ListBox = List1
'Then you can cut and paste.
Private Sub Form_Load()
 'Load our test items.
 Combo1.AddItem "Adam"
 Combo1.AddItem "Bill"
 Combo1.AddItem "Dave"
 Combo1.AddItem "Dick"
 Combo1.AddItem "Neville"
 Combo1.AddItem "Norman"
 Combo1.AddItem "Simon"
 Combo1.AddItem "Steve"
 Combo1.AddItem "Stevie"
 Combo1.AddItem "Tom"
 List1.AddItem "Adam"
 List1.AddItem "Bill"
 List1.AddItem "Dave"
 List1.AddItem "Dick"
 List1.AddItem "Neville"
 List1.AddItem "Norman"
 List1.AddItem "Simon"
 List1.AddItem "Steve"
 List1.AddItem "Stevie"
 List1.AddItem "Tom"
End Sub
Private Sub Text1_Change()
 Dim cmbInd As Long, lstInd As Long
 '0 is the last item in the list not the first
 For cmbInd = (Combo1.ListCount - 1) To 0 Step -1
 If UCase(Left(Combo1.List(cmbInd), Len(Text1.Text))) = UCase(Text1.Text) Then Combo1.ListIndex = cmbInd 'Find and set the selected combo item
 Next
 For lstInd = (List1.ListCount - 1) To 0 Step -1
 If UCase(Left(List1.List(lstInd), Len(Text1.Text))) = UCase(Text1.Text) Then List1.Selected(lstInd) = True 'Find and set the selected list item
 Next
End Sub
```

