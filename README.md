# VBA_UserForm

  The user form designed to help users to validate the input through standarization to help them avoid mistakes.

The form will pop out when the user will click on a comments section.

    If Target.Column = 7 And Target.Row <> 1 And Target.Row <= LastRow Then
        wsData.Range("O1") = Target.Address
        Categories.Show
    End If
    
 
    
    
![userform](https://user-images.githubusercontent.com/86082905/126913170-95550072-941c-4c39-95b2-afa0542d29c5.JPG)

After clicking on a checkbox - a list box will pop out. The user can choose how many comments to provide and how to categorize he comments by using combo boxes (drop down lists)

![33](https://user-images.githubusercontent.com/86082905/126913172-0fe7b5f1-7676-4ab6-b00c-3942aa7271cc.JPG)

The standarized output will be provided after clicing on "Save" button.

![output](https://user-images.githubusercontent.com/86082905/126913173-704178ce-03f2-4b95-8187-5561290ae352.JPG)
