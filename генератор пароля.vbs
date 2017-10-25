option explicit

call main
sub main
msgbox("Генерация пароля")
        Dim Char
        Dim N, Start, GenerateBoundary 
        Const Chars = "abcdefghijklmnopqrstuvxyz0123456789" 
        Randomize                                                                                 
        For N = 1 To 8                                                                        
            Start = CLng(Rnd * (Len(Chars) - 1)) + 1                         
            Char = Mid(Chars, Start, 1)                                                
            If Start Mod 2 Then Char = UCase(Char)                        
            GenerateBoundary = GenerateBoundary & Char           
        Next
msgbox(GenerateBoundary)
end sub