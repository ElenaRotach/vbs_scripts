' ----------------------------------------------------------------------------
' ��������� HeadersFooters � ������ HeaderFooter
' ����������� � Word
' HeadersFooters.vbs
' ----------------------------------------------------------------------------
Option Explicit
 
dim oWord, oDoc, oSel, i, MyText
 
Set oWord = CreateObject("Word.Application")
Set oDoc = oWord.Documents
oDoc.Add()
oWord.Visible = true
Set oSel = oWord.Selection
 
MyText = "����������� � �����. "
 
For i=0 to 40
    oSel.TypeText MyText & MyText & MyText & MyText & MyText & MyText & MyText
    oSel.TypeParagraph
Next
 
With oDoc(1).Sections(1)
    .PageSetup.OddAndEvenPagesHeaderFooter = true                                              ' ��������� ����������� � Word ��� ������ � ��������
    .PageSetup.DifferentFirstPageHeaderFooter = true                                                ' ���������� ���������� � Word ��� ������ ��������

    '-------------------------------------------------------------------------------------------
    ' ������� � ������ ����������� � ����� ��� ������
    '-------------------------------------------------------------------------------------------
    .Headers(3).Range.Text = "��������� � ������1, 2,4,6....."
    .Footers(3).Range.Text = "����� � ������1, 2,4,6....."
    '-------------------------------------------------------------------------------------------

    '-------------------------------------------------------------------------------------------
    ' ������ � ������� ����������� � ����� ��� ��������
    '-------------------------------------------------------------------------------------------
    .Headers(1).Range.Text = "��������� � ��������1, 1,3,5....."
    .Footers(1).Range.Text = "����� � ��������1, 1,3,5....."
    '-------------------------------------------------------------------------------------------

    '-------------------------------------------------------------------------------------------
    ' ������� � ������ ����������� � Word ��� ������ ��������
    '-------------------------------------------------------------------------------------------
    .Headers(2).Range.Text = "���������"
    .Footers(2).Range.Text = "�����"
    '-------------------------------------------------------------------------------------------
End With