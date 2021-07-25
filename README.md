# VBA_UserForm

  The user form designed to help users to validate the input through standarization to help them avoid mistakes.

The form will pop out when the user clicks on the comments section.

    If Target.Column = 7 And Target.Row <> 1 And Target.Row <= LastRow Then
        wsData.Range("O1") = Target.Address
        Categories.Show
    End If
    
 
    
    
![userform](https://user-images.githubusercontent.com/86082905/126913170-95550072-941c-4c39-95b2-afa0542d29c5.JPG)

After ticking a checkbox - a list box pop out. The user can choose how many comments to provide and how to categorize the comments by using combo boxes (drop down lists)



Save button triggers the standarized output.  

